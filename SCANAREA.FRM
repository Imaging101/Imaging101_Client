VERSION 5.00
Object = "{A1FF59A3-0995-11CF-B3E8-00608C82AA8D}#1.6#0"; "PIXEZOCX.OCX"
Begin VB.Form frmScanningArea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Scanning Area"
   ClientHeight    =   3030
   ClientLeft      =   5535
   ClientTop       =   2400
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3030
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin PixezocxLib.PixEzImage PixEzImage1 
      Height          =   2655
      Left            =   3120
      TabIndex        =   12
      Top             =   120
      Width           =   2295
      _Version        =   65542
      _ExtentX        =   4048
      _ExtentY        =   4683
      _StockProps     =   100
      TAG_ENV_TIFFJPEGCOMPRESSION=   6
      TAG_ENV_ANNOTATIONVERSION=   2
      TAG_OPEN_DIR    =   ""
      TAG_OPEN_SCHEMA =   ""
      TAG_OPEN_EXT    =   ""
      TAG_OPEN_ROOT   =   ""
      TAG_OPEN_DETECTSCHEMA=   1
      TAG_OPEN_FILENAMES=   ""
      TAG_WINDOW_CURPAGE=   0
      PIXEZ_SELECT    =   ""
      TAG_BORDER_COLOR_BG=   6579300
      TAG_BORDER_COLOR_ENVFOCUS=   254
      TAG_BORDER_COLOR_WINFOCUS=   16711422
      TAG_BRIGHTNESS  =   128
      TAG_CONTRAST    =   128
      TAG_BLUEBRIGHTNESS=   128
      TAG_BLUECONTRAST=   128
      TAG_GREENBRIGHTNESS=   128
      TAG_GREENCONTRAST=   128
      TAG_REDBRIGHTNESS=   128
      TAG_REDCONTRAST =   128
      TAG_DOC_OPENATTRIBUTE=   66
      TAG_FILLORDER   =   1
      TAG_HFLIP       =   0
      TAG_PAN_HEIGHT  =   0
      TAG_PAN_WIDTH   =   0
      TAG_PAN_XPOS    =   0
      TAG_PAN_YPOS    =   0
      TAG_PAN_SCALING =   4
      TAG_PAN_TITLE   =   "Pan Window"
      TAG_PAN_SHOW    =   0
      TAG_ONE_ACCELMODE=   0
      TAG_ONE_ACTION_CLOCKWISE=   64
      TAG_ONE_ACTION_CTRCLOCKWISE=   64
      TAG_ONE_ACTION_DEFINEREG=   64
      TAG_ONE_ACTION_DEFINEREGASPECT=   64
      TAG_ONE_ACTION_PAN=   64
      TAG_ONE_ACTION_SWITCHTOTREE=   64
      TAG_ONE_ACTION_ZOOMINREG=   64
      TAG_ONE_ACTION_ZOOMINREGASPECT=   64
      TAG_ONE_ACTION_ZOOMOUTCORNER=   64
      TAG_ONE_ACTION_CONTEXTMENU=   64
      TAG_ONE_ACTION_ANNOTATIONITEMTRIGGER=   64
      TAG_ONE_ACTION_ANNOTATIONMODEACTION=   64
      TAG_ONE_ACTION_ANNOTATIONMODEACTIONASP=   64
      TAG_ONE_SCROLLBARS=   0
      TAG_ONE_SETTINGS_RANGE=   0
      TAG_ONE_MOUSEOPTION=   0
      TAG_ORIENTATION =   1
      TAG_OVERSCAN    =   0
      TAG_PHOTOMETRICINTERPRETATION=   0
      TAG_PRINT_COLLATE=   1
      TAG_PRINT_COPIES=   1
      TAG_PRINT_DEVICENO=   0
      TAG_PRINT_DEVNAME1=   ""
      TAG_PRINT_DEVNAME2=   ""
      TAG_PRINT_RANGEMODE=   0
      TAG_PRINT_REGION=   0
      TAG_PRINT_SCALE =   0
      TAG_PRINT_SHOWDLG=   0
      TAG_REGION_COUNT=   0
      TAG_REGION_MODE =   0
      TAG_ROTATION    =   1
      TAG_SCALING     =   4
      TAG_DITHER      =   1
      TAG_VIEWASGRAY  =   0
      TAG_SCALE_X     =   1
      TAG_SCALE_Y     =   1
      TAG_SCAN_ALLOW_TURNOVER=   0
      TAG_SCAN_COLORFORMAT=   8388608
      TAG_SCAN_COMPRESSION=   4
      TAG_SCAN_CURPAGE=   0
      TAG_SCAN_DISPLAYPAGE=   0
      TAG_SCAN_DIR    =   ""
      TAG_SCAN_DUPLEX =   0
      TAG_SCAN_EXT    =   ".tif"
      TAG_SCAN_FILENAME=   "/"
      TAG_SCAN_INSERTMODE=   1
      TAG_SCAN_SCHEMA =   ""
      TAG_SCAN_WARNOVERWRITE=   0
      TAG_SCAN_MULTIPAGE=   1
      TAG_SCAN_USESCHEMA=   0
      TAG_SCAN_MAXPAGES=   -1
      TAG_SCAN_ORIENTATION=   1
      TAG_SCAN_PACK   =   196608
      TAG_SCAN_PRECEDENCE=   1
      TAG_SCAN_ROOT   =   ""
      TAG_SCAN_USELONGNAMES=   0
      TAG_SAVE_MERGEANNOTATIONS=   0
      TAG_SAVE_COLORFORMAT=   8388608
      TAG_SAVE_COMPRESSION=   4
      TAG_SAVE_DIR    =   ""
      TAG_SAVE_EXT    =   ".tif"
      TAG_SAVE_FILENAME=   "/"
      TAG_SAVE_ORIENTATION=   1
      TAG_SAVE_PACK   =   196608
      TAG_SAVE_PRECEDENCE=   1
      TAG_SAVE_RANGESTR=   "1-0"
      TAG_SAVE_ROOT   =   ""
      TAG_SAVE_WARNOVERWRITE=   0
      TAG_SAVE_MULTIPAGE=   0
      TAG_SAVE_USESCHEMA=   0
      TAG_SAVE_USELONGNAMES=   0
      TAG_THRESH_X    =   0
      TAG_THRESH_Y    =   0
      TAG_TREE_COLOR_BG=   8421504
      TAG_TREE_COLOR_NODETEXT=   0
      TAG_TREE_COLOR_NODESELTEXT=   16777215
      TAG_TREE_COLOR_THUMBTEXT=   0
      TAG_TREE_COLOR_THUMBSELTEXT=   16777215
      TAG_TREE_COLOR_LINE=   0
      TAG_TREE_THUMBSTYLE=   528
      TAG_TREE_UIFLAGS=   1280
      TAG_WINDOW_STYLE=   0
      TAG_XPOSITION   =   0
      TAG_YPOSITION   =   0
      TAG_INVERT      =   0
      TAG_SCAN_SCANROT=   0
      TAG_SCAN_AUTOCOLORFORMAT=   0
      TAG_TREE_cxTHUMBNAIL=   34
      TAG_TREE_cxTHUMBCELL=   52
      TAG_TREE_cyTHUMBNAIL=   44
      TAG_TREE_cyTHUMBLINE=   72
      TAG_SAVE_PROGRESSMODE=   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   396
      Left            =   1620
      TabIndex        =   11
      Top             =   2400
      Width           =   1260
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   396
      Left            =   288
      TabIndex        =   10
      Top             =   2400
      Width           =   1260
   End
   Begin VB.TextBox txtLength 
      Height          =   300
      Left            =   1248
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "0"
      Top             =   1824
      Width           =   1740
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   1248
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "0"
      Top             =   1440
      Width           =   1740
   End
   Begin VB.TextBox txtYPosition 
      Height          =   300
      Left            =   1248
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "0"
      Top             =   1056
      Width           =   1740
   End
   Begin VB.TextBox txtXPosition 
      Height          =   300
      Left            =   1248
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "0"
      Top             =   672
      Width           =   1740
   End
   Begin VB.ComboBox cboUnits 
      Height          =   288
      Left            =   1248
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   192
      Width           =   1740
   End
   Begin VB.Label Label5 
      Caption         =   "&Length: "
      Height          =   300
      Left            =   90
      TabIndex        =   8
      Top             =   1824
      Width           =   1068
   End
   Begin VB.Label Label4 
      Caption         =   "&Width:"
      Height          =   300
      Left            =   90
      TabIndex        =   6
      Top             =   1440
      Width           =   1068
   End
   Begin VB.Label Label3 
      Caption         =   "&Y Position:"
      Height          =   300
      Left            =   90
      TabIndex        =   4
      Top             =   1056
      Width           =   1068
   End
   Begin VB.Label Label2 
      Caption         =   "&X Position:"
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Top             =   672
      Width           =   1068
   End
   Begin VB.Label Label1 
      Caption         =   "&Units: "
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Top             =   192
      Width           =   1068
   End
End
Attribute VB_Name = "frmScanningArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const UNIT_PIXELS As Long = 2
Private Const UNIT_MM_THOUSANDS As Long = 1
Private Const UNIT_INCH_THOUSANDS As Long = 0
Private ctrActivePixEzImage As PixEzImage
Private lReturn As Long
Private lCurrentUnit As Long
Private lToUnit As Long
Private firsttime As Integer
Private lSaveUnit As Long
Private bDragging As Integer
Private StartX As Single
Private StartY As Single
Private LastX As Single
Private LastY As Single
Private MaxLength As Long
Private MaxWidth As Long
Private ImageWidth As Long
Private ImageLength As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal bottom As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal rop As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Private Sub cboUnits_Click()
    lToUnit = cboUnits.ItemData(cboUnits.ListIndex)
    If (lToUnit <> lCurrentUnit) Then
        txtXPosition.Text = ConvertUnits(txtXPosition.Text, lToUnit, lCurrentUnit)
        txtYPosition.Text = ConvertUnits(txtYPosition.Text, lToUnit, lCurrentUnit)
        txtWidth.Text = ConvertUnits(txtWidth.Text, lToUnit, lCurrentUnit)
        txtLength.Text = ConvertUnits(txtLength.Text, lToUnit, lCurrentUnit)
        lCurrentUnit = lToUnit
    End If
End Sub


Private Sub cmdCancel_Click()
    cboUnits.ListIndex = lSaveUnit
    Me.Hide
    lReturn = 0
End Sub

Private Function IsDataValid(strValue As String, lCurrentUnit As Long)
       Dim dval As Double
       If (IsNull(strValue) <> True And IsNumeric(strValue) = True) Then
        dval = CDbl(Trim(strValue))
        If (lCurrentUnit = UNIT_PIXELS) Then
            If (InStr(1, strValue, ".") <> 0) Then
                IsDataValid = False
            Exit Function
            End If
        End If
        If (dval >= 0) Then
            IsDataValid = True
        Else
            IsDataValid = False
        End If
       Else
        IsDataValid = False
       End If

End Function

Private Sub cmdOK_Click()
    Dim Top, Left, width, length As Long
    
    lCurrentUnit = cboUnits.ItemData(cboUnits.ListIndex)
    
    ' Numerical Valdiation
    
    If (IsDataValid(txtXPosition.Text, lCurrentUnit) = False) Then
        MsgBox EW2W("X Position Is Not Valid")
        txtXPosition.Text = ConvertUnits(ctrActivePixEzImage.ScanLeft, lToUnit, UNIT_PIXELS)
        txtXPosition.SetFocus
        Exit Sub
    End If
    
    If (IsDataValid(txtYPosition.Text, lCurrentUnit) = False) Then
        MsgBox EW2W("Y Position Is Not Valid")
        txtYPosition.Text = ConvertUnits(ctrActivePixEzImage.ScanTop, lToUnit, UNIT_PIXELS)
        txtYPosition.SetFocus
        Exit Sub
    End If
    
    If (IsDataValid(txtWidth.Text, lCurrentUnit) = False) Then
        MsgBox EW2W("Width Is Not Valid")
        txtWidth.Text = ConvertUnits(ctrActivePixEzImage.ScanWidth, lToUnit, UNIT_PIXELS)
        txtWidth.SetFocus
        Exit Sub
    End If
    
    If (IsDataValid(txtLength.Text, lCurrentUnit) = False) Then
        MsgBox EW2W("Length Is Not Valid")
        txtLength.Text = ConvertUnits(ctrActivePixEzImage.ScanLength, lToUnit, UNIT_PIXELS)
        txtLength.SetFocus
        Exit Sub
    End If
    
    ' save the user-entered values
    Top = ConvertUnits(txtYPosition.Text, UNIT_PIXELS, lCurrentUnit)
    Left = ConvertUnits(txtXPosition.Text, UNIT_PIXELS, lCurrentUnit)
    width = ConvertUnits(txtWidth.Text, UNIT_PIXELS, lCurrentUnit)
    length = ConvertUnits(txtLength.Text, UNIT_PIXELS, lCurrentUnit)
    
    'set the values into the image
    ctrActivePixEzImage.ScanTop = Top
    ctrActivePixEzImage.ScanLeft = Left
    ctrActivePixEzImage.ScanWidth = width
    ctrActivePixEzImage.ScanLength = length
    
    'check the rebound. If they're not the same , the user has exceeded the the limits someway.
    'Display the current values and dont close the form.
    If Top > ctrActivePixEzImage.ScanTop Or Left > ctrActivePixEzImage.ScanLeft Or width > ctrActivePixEzImage.ScanWidth Or length > ctrActivePixEzImage.ScanLength Then
        MsgBox EW2W("Out of Bounds")
        txtXPosition.Text = ConvertUnits(ctrActivePixEzImage.ScanLeft, lToUnit, UNIT_PIXELS)
        txtYPosition.Text = ConvertUnits(ctrActivePixEzImage.ScanTop, lToUnit, UNIT_PIXELS)
        txtWidth.Text = ConvertUnits(ctrActivePixEzImage.ScanWidth, lToUnit, UNIT_PIXELS)
        txtLength.Text = ConvertUnits(ctrActivePixEzImage.ScanLength, lToUnit, UNIT_PIXELS)
        Exit Sub
    End If
    
    Me.Hide
    lReturn = 1
End Sub

Private Sub Form_Activate()
    txtXPosition.SetFocus
End Sub

Private Sub Form_Load()
   PixLocalizeForm Me
   ' center form
   '
    PixLocalizeForm Me
   cboUnits.Clear
   cboUnits.AddItem EW2W("Inches")
   cboUnits.ItemData(cboUnits.NewIndex) = UNIT_INCH_THOUSANDS
   cboUnits.AddItem EW2W("Millimeters")
   cboUnits.ItemData(cboUnits.NewIndex) = UNIT_MM_THOUSANDS
   cboUnits.AddItem EW2W("Pixels")
   cboUnits.ItemData(cboUnits.NewIndex) = UNIT_PIXELS
    
   'Me.Move (Screen.width - Me.width) / 2, _
        (Screen.Height - Me.Height) / 2
   'cboUnits_Click
End Sub
Public Function StartForm(ctrPixEzImage As PixEzImage) As Long
    Set ctrActivePixEzImage = ctrPixEzImage
    
    Dim t As Long, l As Long, w As Long, h As Long
    With ctrPixEzImage
    h = .ScanLength
    w = .ScanWidth
    t = .ScanTop
    l = .ScanLeft
    
    .ScanLeft = 0
    .ScanTop = 0
    .ScanWidth = 32000
    .ScanLength = 32000
    
    MaxWidth = .ScanWidth
    MaxLength = .ScanLength
    ImageWidth = PixEzImage1.width
    ImageLength = PixEzImage1.Height
        
    .ScanLeft = l
    .ScanTop = t
    .ScanWidth = w
    .ScanLength = h
    End With
    
    If ctrActivePixEzImage.PageCount Then
        PixEzImage1.DocHandle = ctrActivePixEzImage.DocHandle
       
        Dim lPct As Long, wPct As Long
        If PixEzImage1.width / MaxWidth < PixEzImage1.Height / MaxLength Then
            PixEzImage1.ScaleRatio.Set PixEzImage1.width * 10, MaxWidth * 149
            ImageLength = PixEzImage1.width * MaxLength / MaxWidth
        Else
            PixEzImage1.ScaleRatio.Set PixEzImage1.Height * 10, MaxLength * 149
            ImageWidth = PixEzImage1.Height * MaxWidth / MaxLength
        End If
    Else
    End If
    'width = MaxWidth
    
    On Error Resume Next
    lCurrentUnit = UNIT_PIXELS
    If Not firsttime Then
        lToUnit = UNIT_INCH_THOUSANDS
    End If
    firsttime = True
    ReloadData
    'txtYPosition.Text = ctrActivePixEzImage.ScanTop
    'txtXPosition.Text = ctrActivePixEzImage.ScanLeft
    'txtWidth.Text = ctrActivePixEzImage.ScanWidth
    'txtLength.Text = ctrActivePixEzImage.ScanLength

    'txtXPosition.Text = ConvertUnits(ctrActivePixEzImage.ScanTop, lToUnit, lCurrentUnit)
    'txtYPosition.Text = ConvertUnits(ctrActivePixEzImage.ScanLeft, lToUnit, lCurrentUnit)
    'txtWidth.Text = ConvertUnits(ctrActivePixEzImage.ScanWidth, lToUnit, lCurrentUnit)
    'txtLength.Text = ConvertUnits(ctrActivePixEzImage.ScanLength, lToUnit, lCurrentUnit)
    'lCurrentUnit = lToUnit
    cboUnits.ListIndex = lToUnit
    cboUnits_Click
    lSaveUnit = cboUnits.ListIndex
    Me.Show 1
    StartForm = lReturn
    Me.Hide
End Function

Private Sub PixEzImage1_DblClick()
    txtXPosition.Text = ConvertUnits(CStr(0), lToUnit, UNIT_PIXELS)
    txtYPosition.Text = ConvertUnits(CStr(0), lToUnit, UNIT_PIXELS)
    txtWidth.Text = ConvertUnits(CStr(MaxWidth), lToUnit, UNIT_PIXELS)
    txtLength.Text = ConvertUnits(CStr(MaxLength), lToUnit, UNIT_PIXELS)
    
End Sub

Private Sub PixEzImage1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetCapture(PixEzImage1.hwnd)
PixEzImage1.Refresh
bDragging = 1
StartX = X
StartY = Y
LastX = X
LastY = Y
Call DrawBox(StartX, StartY, LastX, LastY)
End Sub

Private Function Min(ByVal a, ByVal b)
If a < b Then
    Min = a
Else
    Min = b
End If
End Function

Private Function Max(ByVal a, ByVal b)
If a > b Then
    Max = a
Else
    Max = b
End If
End Function

Private Sub PixEzImage1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (bDragging) Then
    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single
    
    Call DrawBox(StartX, StartY, LastX, LastY)
    Call DrawBox(StartX, StartY, X, Y)
    LastX = X
    LastY = Y
        
    x1 = Min(1, Max(0, Min(X, StartX)) / ImageWidth) * MaxWidth
    x2 = Min(1, Max(0, Max(X, StartX)) / ImageWidth) * MaxWidth
    y1 = Min(1, Max(0, Min(Y, StartY)) / ImageLength) * MaxLength
    y2 = Min(1, Max(0, Max(Y, StartY)) / ImageLength) * MaxLength
    
    txtXPosition.Text = ConvertUnits(CStr(x1), lToUnit, UNIT_PIXELS)
    txtYPosition.Text = ConvertUnits(CStr(y1), lToUnit, UNIT_PIXELS)
    txtWidth.Text = ConvertUnits(CStr(x2 - x1), lToUnit, UNIT_PIXELS)
    txtLength.Text = ConvertUnits(CStr(y2 - y1), lToUnit, UNIT_PIXELS)
    End If
    
End Sub



Private Sub PixEzImage1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bDragging Then
    Call PixEzImage1_MouseMove(Button, Shift, X, Y)
    bDragging = 0
    End If
ReleaseCapture
End Sub

Sub DrawBox(ByVal cX1 As Single, ByVal cY1 As Single, ByVal cX2 As Single, ByVal cY2 As Single)
Dim hdc As Long
Dim x1 As Single, x2 As Single, y1 As Single, y2 As Single

Const LineWidth = 3
Const factor = 15

hdc = GetDC(PixEzImage1.hwnd)
Call SetROP2(hdc, 7)

x1 = Min(cX1, cX2) / factor
x2 = Max(cX1, cX2) / factor
y1 = Min(cY1, cY2) / factor
y2 = Max(cY1, cY2) / factor

Call Rectangle(hdc, x1, y1, x1 + LineWidth, y2)
Call Rectangle(hdc, x1 + LineWidth - 2, y2, x2, y2 - LineWidth)
Call Rectangle(hdc, x1 + LineWidth - 2, y1, x2, y1 + LineWidth)
Call Rectangle(hdc, x2 - 2, y1, x2 + LineWidth - 2, y2)

Call ReleaseDC(PixEzImage1.hwnd, hdc)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ctrActivePixEzImage = Nothing
End Sub

Private Function ConvertUnits(valOld As String, Unitto As Long, Unitfrom As Long)
    Dim dvalNew  As Double
    Dim dvalOld  As Double
    Dim OldValue As String
    On Error GoTo ERROR_HANDLER
    OldValue = Trim(valOld)
    If (OldValue = Null Or OldValue = "" Or OldValue = " ") Then
        ConvertUnits = ""
        Exit Function
    End If
    dvalOld = CDbl(OldValue)
    dvalNew = dvalOld
    If (Unitfrom <> Unitto) Then
        ' Convert Unitto UNIT_INCH_THOUSANDS
        Select Case Unitfrom
            Case UNIT_PIXELS
                  '  dvalNew = dvalOld * 1000 * res->Denom / res->Num
                  dvalNew = dvalOld / ctrActivePixEzImage.ScanResolution
            Case UNIT_MM_THOUSANDS
                    dvalNew = (dvalOld * 10) / 254
            Case UNIT_INCH_THOUSANDS
                    dvalNew = dvalOld
        End Select
        
        ' Validation
        If (dvalNew > 100000) Then
             dvalNew = 100000
        End If
      
            ' Convert Unitfrom UNIT_INCH_THOUSANDS
        Select Case Unitto
            Case UNIT_PIXELS
                 dvalNew = (dvalNew * ctrActivePixEzImage.ScanResolution)
            Case UNIT_MM_THOUSANDS
                    dvalNew = (dvalNew * 254) / 10
            Case UNIT_INCH_THOUSANDS
                    dvalNew = dvalNew
        End Select
    End If
    ' Formating
    Select Case Unitto
        Case UNIT_PIXELS
             ConvertUnits = Format(dvalNew, "######0")
        Case UNIT_MM_THOUSANDS
             ConvertUnits = Format(dvalNew, "######0.000")
         Case UNIT_INCH_THOUSANDS
             ConvertUnits = Format(dvalNew, "#####0.000")
    End Select
    Exit Function
ERROR_HANDLER:
   ConvertUnits = ""
End Function

Private Sub ReloadData()
    txtYPosition.Text = ctrActivePixEzImage.ScanTop
    txtXPosition.Text = ctrActivePixEzImage.ScanLeft
    txtWidth.Text = ctrActivePixEzImage.ScanWidth
    txtLength.Text = ctrActivePixEzImage.ScanLength
End Sub

Private Sub txtLength_GotFocus()
    txtLength.SelStart = 0  ' set selection start and
    txtLength.SelLength = Len(txtLength)   ' set selection length.
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = 0  ' set selection start and
    txtWidth.SelLength = Len(txtWidth)   ' set selection length.
End Sub

Private Sub txtXPosition_GotFocus()
    txtXPosition.SelStart = 0  ' set selection start and
    txtXPosition.SelLength = Len(txtXPosition)   ' set selection length.
End Sub

Private Sub txtYPosition_GotFocus()
    txtYPosition.SelStart = 0  ' set selection start and
    txtYPosition.SelLength = Len(txtYPosition)   ' set selection length.
End Sub

