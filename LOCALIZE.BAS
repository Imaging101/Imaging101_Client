Attribute VB_Name = "localize"
'****************************************************************************
'   "$Workfile: Localize.BAS $"
'
'   Copyright (c) 1998-2003 Captiva Software
'   All Rights Reserved.
'   Proprietary to Input Software
'   Your use is governed by your license agreement.
'
'****************************************************************************

Option Explicit

' $Workfile: Localize.BAS $ $Date: 5/21/03 9:56a $ $LockStatus: $ '

' ---------------------------------------------------------------------------
  Const SourceFileVersion = "&(#)$Workfile: Localize.BAS $ v$Revision: 69 $ $Date: 5/21/03 9:56a $"
' ---------------------------------------------------------------------------

Declare Sub PixLocalizeSetLocale Lib "PixLocn.dll" (ByVal lLocale As Long)
Declare Function PixLocalizeGetLocale Lib "PixLocn.dll" () As Long
Declare Sub PixLocalizeAddFile Lib "PixLocn.dll" (ByVal lpszFile As String)
Declare Function PixLocalizeLookupVariant Lib "PixLocn.dll" (ByVal lpIn As Variant, lpOut As Variant) As Long
Declare Sub PixLocalizeGetStringSize Lib "PixLocn.dll" (ByVal hdc As Long, ByVal lpIn As Variant, sz As size)

Type size
   cx As Long
   cy As Long
End Type
   
Declare Sub PixLocalizeScaleSize Lib "PixLocn.dll" (ByVal szFace As String, ByVal wFontSize As Integer, lpSize As size)
Declare Sub PixLocalizeGetCurSizeIFont Lib "PixLocn.dll" (ByVal Font As IFontDisp, lpSize As size)

Declare Sub PixLocalizeSetLanguage Lib "PixLocn.dll" (ByVal dwLanguage As Long)
Declare Sub PixLocalizeGetLanguage Lib "PixLocn.dll" (dwLanguage As Long)
Declare Function GetProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Declare Function GetProfileInt Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer

Declare Sub GetTextExtentPoint32A Lib "GDI32.DLL" (ByVal hdc As Long, ByVal lpStr As String, ByVal nChars As Long, sz As size)

Private Const PIXEL_MARGIN = 0

Public Const LOC_SNAP_MAX_RIGHT = 0
Public Const LOC_SNAP_MAX_BOTTOM = 1
Public Const LOC_SNAP_MAX_TOP = 2
Public Const LOC_SNAP_MAX_LEFT = 3

Public Const LANGUAGE_ENGLISH = 1033
Public Const LANGUAGE_ITALIAN = 1040
Public Const LANGUAGE_GERMAN = 1031
Public Const LANGUAGE_FRENCH = 1036
Public Const LANGUAGE_SPANISH = 1034
Public Const LANGUAGE_PORTUGESE = 2070
Public Const LANGUAGE_SWEDISH = 1053
Public Const LANGUAGE_JAPANESE = 1041

Public Type PixLocState
    xScale As Integer
    yScale As Integer
    doFont As Boolean
    handleTabs As Boolean
    numTabAnchors As Integer
    doRed As Boolean
End Type

Function PixLocalizeFont(fnt As StdFont) As Boolean
    On Error Resume Next

    Dim fntSys As StdFont
    Dim fntNew As String
    Dim fntDesc As String
    Dim fntParts() As String

    fntDesc = fnt.name & "/" & fnt.size
    fntNew = EW2W(fntDesc)
    ' If no translation, try to find a scaling amount
    If fntNew = fntDesc Then
        fntDesc = fnt.name & "/Scaled"
        fntNew = EW2W(fntDesc)
        If fntNew <> fntDesc Then
            Dim size As Currency
            fntParts = Split(fntNew, "/", 2)
            fnt.name = fntParts(0)
            fnt.size = fnt.size * Int(fntParts(1)) / 100
            Set fntSys = New StdFont
            fntSys.name = "System"
            fnt.Charset = fntSys.Charset
            PixLocalizeFont = True
        Else
            PixLocalizeFont = False
        End If
    Else
        fntParts = Split(fntNew, "/", 2)
        fnt.name = fntParts(0)
        fnt.size = fntParts(1)
        Set fntSys = New StdFont
        fntSys.name = "System"
        fnt.Charset = fntSys.Charset
        PixLocalizeFont = True
    End If

    Err.Clear
End Function

Sub PixLocalizeForm(frm As Form)
    On Error Resume Next

    Dim ctl As Control
    Dim obj As Object
    Dim strType As String
    Dim sz1 As size
    Dim sz2 As size
    Dim State As PixLocState
    
    ' skip everything if we're in english.
    If (PixLocalizeGetLocale() = LANGUAGE_ENGLISH) Then
        Exit Sub
    End If
    
    InitializeState State
    
    State.doRed = GetSetting(App.Title, "Settings", "LongStringsRed", 0)
    
    PixLocalizeGetFontInfo frm.Font, State

    ' Determine if this is a tabbed control with anchors
    If State.doFont Then
        For Each ctl In frm.Controls
            strType = TypeName(ctl)
            If strType = "SSTab" Then
                State.handleTabs = True
            End If
            If strType = "Frame" And Left(ctl.name, 12) = "PixTabAnchor" Then
                State.numTabAnchors = State.numTabAnchors + 1
            End If
        Next
    End If
    
    ' Some helpful? debugging output
    If State.doRed Then
        Dim fd As Integer
        fd = FreeFile
        If fd <> 0 Then
            Open "c:\loc.log" For Append Lock Write As fd
            Print #fd, "[" & frm.Caption & " - " & PixLocalizeGetLocale() & "]"
            Close fd
        End If
    End If
    
    ' Set the form's caption
    frm.Caption = EW2W(frm.Caption)
    
    ' Scale the form
    If State.doFont Then
        frm.width = frm.width * State.xScale / 1000
        frm.Height = frm.Height * State.yScale / 1000
    End If
    
    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        PixLocalizeOneControl frm, ctl, State
    Next
    
    ' Clear any errors
    Err.Clear
End Sub

Private Sub PixLocalizeGetFontInfo(Font As IFontDisp, State As PixLocState)
    On Error Resume Next
    
    ' get the current font dimensions
    Dim sz1 As size
    Dim sz2 As size
    
    PixLocalizeGetCurSizeIFont Font, sz1
    If PixLocalizeFont(Font) Then
        ' get the new dimensions
        PixLocalizeGetCurSizeIFont Font, sz2
        ' compute the scaling
        State.xScale = sz2.cx * 1000 / sz1.cx
        State.yScale = sz2.cy * 1000 / sz1.cy
        State.doFont = True
    Else
        State.doFont = False
    End If

End Sub

Function PixLocalizeGetUserControlState(Controls As Object, Font As IFontDisp) As PixLocState
    On Error Resume Next
    
    Dim State As PixLocState
    InitializeState State
    
    State.doRed = GetSetting(App.Title, "Settings", "LongStringsRed", 0)
    
    PixLocalizeGetFontInfo Font, State
    
    ' Determine if this is a tabbed control with anchors
    If State.doFont Then
        Dim ctl As Control
        Dim strType As String
        Dim i As Integer
        For Each ctl In Controls
            strType = TypeName(ctl)
            If strType = "SSTab" Then
                State.handleTabs = True
            End If
            If strType = "Frame" And Left(ctl.name, 12) = "PixTabAnchor" Then
                State.numTabAnchors = State.numTabAnchors + 1
            End If
        Next
    End If
    
    PixLocalizeGetUserControlState = State

End Function

Private Function PixLocalizeNoSizeChanges(Control As Object) As Boolean
    If TypeName(Control) = "CommandButton" Then
        If Control.Style = 1 Then
            PixLocalizeNoSizeChanges = True
        End If
    End If
End Function

Sub PixLocalizeUserControl(Parent As Object, Controls As Object, Font As IFontDisp)
    On Error Resume Next
    
    ' skip everything if we're in english.
    If (PixLocalizeGetLocale() = LANGUAGE_ENGLISH) Then
        Exit Sub
    End If
        
    Dim State As PixLocState
    State = PixLocalizeGetUserControlState(Controls, Font)
    
    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    Dim ctl As Control
    For Each ctl In Controls
        PixLocalizeOneControl Parent, ctl, State
    Next
    
    ' Clear any errors
    Err.Clear
End Sub

Sub PixLocalizeOneControl(Parent As Object, ctl As Control, State As PixLocState)
    On Error Resume Next
    Dim obj As Object
    Dim strType As String
    strType = TypeName(ctl)
    
    ' Do moving and sizing and font changes
    Select Case strType
        Case "PictureBox"
            ' Don't Resize PictureBox.
        Case "Line"
            If State.doFont Then
                ctl.x1 = ctl.x1 * State.xScale / 1000
                ctl.y1 = ctl.y1 * State.yScale / 1000
                ctl.x2 = ctl.x2 * State.xScale / 1000
                ctl.y2 = ctl.y2 * State.yScale / 1000
            End If
        Case Else
            Dim newHeight As Integer
            If State.doFont And Left(ctl.name, 12) <> "PixTabAnchor" Then
                Dim leftAnchor As Single
                Dim topAnchor As Single
                leftAnchor = 0
                topAnchor = 0

                If State.handleTabs And State.numTabAnchors > 0 And Left(ctl.name, 12) <> "PixTabAnchor" Then
                    ComputeAnchor Parent, ctl, leftAnchor, topAnchor
                End If
                If leftAnchor <> 0 And State.xScale <> 1000 Then
                    Dim saveL As Single
                    saveL = ctl.Left
                    Dim saveT As Single
                    saveT = ctl.Top
                    ctl.Left = leftAnchor + ((ctl.Left - leftAnchor) * State.xScale / 1000)
                Else
                    ctl.Left = ctl.Left * State.xScale / 1000
                End If
                ctl.Top = ctl.Top * State.yScale / 1000
                If Not PixLocalizeNoSizeChanges(ctl) Then ctl.width = ctl.width * State.xScale / 1000
                
                If Not PixLocalizeNoSizeChanges(ctl) Then
                    ' Don't set List Box Height until after setting font
                    If SetHeightAfter(strType) Then
                        newHeight = ctl.Height * State.yScale / 1000
                    Else
                        ctl.Height = ctl.Height * State.yScale / 1000
                    End If
                End If
                
                If strType = "SSTab" Then
                    ctl.TabHeight = ctl.TabHeight * State.yScale / 1000
                End If
            End If
            ' SSTabs change font when the form font changes
            If strType <> "SSTab" Then
                PixLocalizeFont ctl.Font
            End If
            
            If Not PixLocalizeNoSizeChanges(ctl) Then
                ' Set List Box Height after setting font
                If SetHeightAfter(strType) Then
                    ctl.Height = newHeight
                End If
            End If
    End Select
    
    ' Do String localization
    Select Case strType
        Case "Menu"
            ctl.Caption = EW2W(ctl.Caption)
        Case "Label"
            ctl.Caption = EW2W(ctl.Caption)
            MakeRed State.doRed, Parent.hdc, ctl, ctl.Caption
        Case "TabStrip"
            For Each obj In ctl.Tabs
                obj.Caption = EW2W(obj.Caption)
                obj.ToolTipText = EW2W(obj.ToolTipText)
            Next
        Case "Toolbar"
            For Each obj In ctl.Buttons
                obj.ToolTipText = EW2W(obj.ToolTipText)
                obj.Caption = EW2W(obj.Caption)
                obj.Description = EW2W(obj.Description)
            Next
        Case "ListView"
            For Each obj In ctl.ColumnHeaders
                obj.Text = EW2W(obj.Text)
            Next
        Case "TextBox"
            ctl.Text = EW2W(ctl.Text)
            MakeRed State.doRed, Parent.hdc, ctl, ctl.Text
        Case "PixEzSaveControl"
            ' skip if it's the save control range edit box...
            If ctl.ControlType <> 10 Then
                ctl.Caption = EW2W(ctl.Caption)
                ctl.ToolTipText = EW2W(ctl.ToolTipText)
            End If
            MakeRed State.doRed, Parent.hdc, ctl, ctl.Caption
        Case Else
            ctl.Caption = EW2W(ctl.Caption)
            MakeRed State.doRed, Parent.hdc, ctl, ctl.Caption
            ctl.ToolTipText = EW2W(ctl.ToolTipText)
    End Select
    
    ' Clear any errors
    Err.Clear
End Sub
Function SetHeightAfter(t As String) As Boolean
    SetHeightAfter = False
    If t = "ListBox" Or t = "FileListBox" Or t = "DirListBox" Then
        SetHeightAfter = True
    End If
End Function
Function InRect(ByVal frm As Form, ByVal ctlItem As Control, ByVal ctlAnchor As Control) As Boolean
    If (ctlItem.Left >= ctlAnchor.Left And ctlItem.Left <= ctlAnchor.Left + frm.width And ctlItem.Top >= ctlAnchor.Top And ctlItem.Top <= ctlAnchor.Top + frm.Height) Then
        InRect = True
    Else
        InRect = False
    End If
End Function
Sub ComputeAnchor(ByVal frm As Form, ByVal ctlRef As Control, ByRef leftAnchor As Single, ByRef topAnchor As Single)
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeName(ctl) = "Frame" And Left(ctl.name, 12) = "PixTabAnchor" Then
            If InRect(frm, ctlRef, ctl) Then
                leftAnchor = ctl.Left
                topAnchor = ctl.Top
                Return
            End If
        End If
    Next
End Sub
Sub MakeRed(ByVal doRed As Integer, ByVal hdc As Long, ByVal ctl As Control, ByVal str As String)
    Dim itemsize As size
    If doRed = 1 Then
        'GetTextExtentPoint32A hdc, str, Len(str), itemsize
        Call PixLocalizeGetStringSize(hdc, str, itemsize)
        itemsize.cx = itemsize.cx * Screen.TwipsPerPixelX
        If itemsize.cx > ctl.width Then
            ctl.ForeColor = RGB(255, 0, 0)
            Dim fd As Integer
            fd = FreeFile
            If fd <> 0 Then
                Open "c:\loc.log" For Append Lock Write As fd
                Print #fd, "  " & str & "(" & (itemsize.cx / Screen.TwipsPerPixelX) & ") width: " & (ctl.width / Screen.TwipsPerPixelX)
                Close fd
            End If
        End If
    End If
End Sub
Public Sub InitPixLocalize(Optional ByVal lLocale As Variant, Optional ByVal strfile As Variant)
        Dim lUseLocale As Long
        Dim lLen As Long
        Dim strTemp As String * 255
        Dim strUseFile As String
        
        On Error Resume Next
        
        ' Get locale and translation file, if applicable
        If IsMissing(lLocale) Then
            lUseLocale = 0
        Else
            lUseLocale = lLocale
        End If
        
        If lUseLocale <> 0 Then
            Call PixLocalizeSetLocale(lUseLocale)
            ' Tell PixDflt what our locale is
            PixLocalizeSetLanguage lUseLocale
        End If
        
        If (IsMissing(strfile) = False) Then
            strUseFile = strfile
        End If
        If strUseFile <> "" Then
            Call PixLocalizeAddFile(App.path & "\" & strUseFile)
        End If
        
End Sub

Public Sub UninitPixLocalize()
End Sub

Public Function EW2W(ByVal s As String) As String
    Dim lResult As Long
    Dim strOut As String
    
    lResult = PixLocalizeLookupVariant(s, strOut)
    If lResult = 0 Then
        EW2W = strOut
    Else
        EW2W = s
    End If
End Function

Public Function EW2EW(ByVal s As String) As String
    EW2EW = s
End Function

Function VarSub(s As String, vars) As String
    Dim i As Integer
    Dim ans As String
    ans = s
    For i = UBound(vars) To LBound(vars) Step -1
        ans = Replace(ans, "%" & (i - LBound(vars) + 1), vars(i))
    Next i
    VarSub = ans
End Function



Public Sub PixLocalizeManuallySizeControlXY(hdc As Long, ctl1 As Control, lMinWidth As Long, lMaxWidth As Long, Optional lDeltaX As Long, Optional lDeltaY As Long)
Dim itemsize As size
Dim lBeforeWidth As Long
Dim lBeforeHeight As Long

    PixLocalizeGetStringSize hdc, ctl1.Caption, itemsize
    itemsize.cx = itemsize.cx * Screen.TwipsPerPixelX
    itemsize.cy = itemsize.cy * Screen.TwipsPerPixelY
    lBeforeWidth = (getControlBasicWidthInTwips(ctl1) + itemsize.cx) - ctl1.width
    lBeforeHeight = (getControlBasicHeightInTwips(ctl1) + itemsize.cy) - ctl1.Height
    ctl1.width = getControlBasicWidthInTwips(ctl1) + itemsize.cx
    
        
    If ctl1.width > lMaxWidth Then
        'the height should increase for every time the width does not fit.
        ctl1.Height = getControlBasicHeightInTwips(ctl1) + itemsize.cy
        While ctl1.width > lMaxWidth
            ctl1.width = ctl1.width - lMaxWidth
            ctl1.Height = ctl1.Height + itemsize.cy
        Wend
    End If
    
    lDeltaX = lDeltaX + (ctl1.width - lBeforeWidth)
    If lMinWidth > 0 And ctl1.width < lMinWidth Then
        ctl1.width = lMinWidth
    End If
End Sub


Public Sub PixLocalizeManuallySizeControl(hdc As Long, ctl1 As Control, lMinWidth As Long, Optional lDelta As Long)
Dim itemsize As size
Dim lBeforeWidth As Long

    PixLocalizeGetStringSize hdc, ctl1.Caption, itemsize
    itemsize.cx = itemsize.cx * Screen.TwipsPerPixelX
    lBeforeWidth = (getControlBasicWidthInTwips(ctl1) + itemsize.cx) - ctl1.width
    ctl1.width = getControlBasicWidthInTwips(ctl1) + itemsize.cx
    lDelta = lDelta + (ctl1.width - lBeforeWidth)
    If lMinWidth > 0 And ctl1.width < lMinWidth Then
        ctl1.width = lMinWidth
    End If
End Sub


Public Function PixLocalizeGetProposedControlSize(hdc As Long, ctl1 As Control, lMinWidth As Long) As Long
Dim itemsize As size

    PixLocalizeGetStringSize hdc, ctl1.Caption, itemsize
    itemsize.cx = itemsize.cx * Screen.TwipsPerPixelX
    PixLocalizeGetProposedControlSize = getControlBasicWidthInTwips(ctl1) + itemsize.cx
    If lMinWidth > 0 And PixLocalizeGetProposedControlSize < lMinWidth Then
        PixLocalizeGetProposedControlSize = lMinWidth
    End If
End Function


Public Sub PixLocalizeControl(hdc As Long, ctl1 As Control, ctl2 As Control, lExtraPixelMargin As Long, lMinWidth As Long)
Dim itemsize As size

    PixLocalizeGetStringSize hdc, ctl1.Caption, itemsize
    itemsize.cx = itemsize.cx * Screen.TwipsPerPixelX
    
    ctl1.width = getControlBasicWidthInTwips(ctl1) + itemsize.cx
    If lMinWidth > 0 And ctl1.width < lMinWidth Then
        ctl1.width = lMinWidth
    End If
    ctl2.Left = ctl1.Left + ctl1.width + ((PIXEL_MARGIN + lExtraPixelMargin) * Screen.TwipsPerPixelX)
End Sub


Public Function getControlBasicWidthInTwips(ByVal ctl As Control) As Long
    getControlBasicWidthInTwips = 0
    If TypeOf ctl Is OptionButton Then
        getControlBasicWidthInTwips = 32 * Screen.TwipsPerPixelX
        Exit Function
    End If
    If TypeOf ctl Is CheckBox Then
        getControlBasicWidthInTwips = 30 * Screen.TwipsPerPixelX
        Exit Function
    End If
    If TypeOf ctl Is CommandButton Then
        getControlBasicWidthInTwips = 30 * Screen.TwipsPerPixelX
        Exit Function
    End If
End Function


Public Function getControlBasicHeightInTwips(ByVal ctl As Control) As Long
    getControlBasicHeightInTwips = 0
    If TypeOf ctl Is OptionButton Then
        Exit Function
    End If
    If TypeOf ctl Is CheckBox Then
        getControlBasicHeightInTwips = 0 * Screen.TwipsPerPixelY
        Exit Function
    End If
End Function


Public Function PixLocalizeControlsFitWidth(plWidth As Long, ctls) As Boolean
Dim i As Integer
                    
    For i = LBound(ctls) To UBound(ctls)
        ctls(i).width = plWidth - ctls(i).Left
    Next i
End Function


Public Function PixLocalizeControlsFitMaxWidth(ctls, lMaxRight As Long) As Boolean
Dim i As Integer
Dim lMax As Long

    lMax = 0
    For i = LBound(ctls) To UBound(ctls)
        If ctls(i).width > lMax Then lMax = ctls(i).width
    Next i
    For i = LBound(ctls) To UBound(ctls)
        ctls(i).width = lMax
    Next i
    If lMaxRight > 0 Then
        For i = LBound(ctls) To UBound(ctls)
            ctls(i).Left = lMaxRight - ctls(i).width
        Next i
    End If
End Function

Public Function PixLocalizeSnapControlPositions(ByVal undoRed As Integer, pbLeftJustify As Boolean, piDirection As Integer, ctls) As Boolean
Dim i As Integer
Dim lMax As Long
    
    If undoRed = 1 Then
        For i = LBound(ctls) To UBound(ctls)
            ctls(i).ForeColor = RGB(0, 0, 0)
        Next i
    End If
    
    Select Case LOC_SNAP_MAX_RIGHT
        Case LOC_SNAP_MAX_TOP
                lMax = 9999999
                For i = LBound(ctls) To UBound(ctls)
                    If ctls(i).Top > lMax Then lMax = ctls(i).Top
                Next i
                For i = LBound(ctls) To UBound(ctls)
                    ctls(i).Top = lMax
                Next i
        
        Case LOC_SNAP_MAX_BOTTOM
                lMax = 0
                For i = LBound(ctls) To UBound(ctls)
                    If ctls(i).Top > lMax Then lMax = ctls(i).Top
                Next i
                For i = LBound(ctls) To UBound(ctls)
                    ctls(i).Top = lMax
                Next i
        
        Case LOC_SNAP_MAX_RIGHT
                If pbLeftJustify Then
                    lMax = 0
                    For i = LBound(ctls) To UBound(ctls)
                        If ctls(i).Left > lMax Then lMax = ctls(i).Left
                    Next i
                    For i = LBound(ctls) To UBound(ctls)
                        ctls(i).Left = lMax
                    Next i
                Else
                    lMax = 0
                    For i = LBound(ctls) To UBound(ctls)
                        If ctls(i).Left + ctls(i).width > lMax Then lMax = ctls(i).Left + ctls(i).width
                    Next i
                    For i = LBound(ctls) To UBound(ctls)
                        ctls(i).Left = lMax - ctls(i).width
                    Next i
                End If
            
        Case LOC_SNAP_MAX_LEFT
                lMax = 9999999
                For i = LBound(ctls) To UBound(ctls)
                    If ctls(i).Left > lMax Then lMax = ctls(i).Left
                Next i
                For i = LBound(ctls) To UBound(ctls)
                    ctls(i).Left = lMax
                Next i
    End Select
End Function
Function PixLocalizeIsChildOf(ctl As Control, ctlParent As Control) As Boolean
    PixLocalizeIsChildOf = False
    On Error GoTo NoContainer
    If ctl.Container.name = ctlParent.name Then
        PixLocalizeIsChildOf = True
    End If
    Exit Function
NoContainer:
End Function
Sub PixLocalizeFixTabControl(frm As Form, ctl As Control, ctlsBelow, Optional ByVal DesiredWidth As Integer = -1)
    Dim i As Integer
    Dim rowsNeeded As Integer
    Dim widthSoFar As Long
    Dim minTabsPerRow As Integer
    Dim tabsThisRow As Integer
    Dim border As Integer
    
    If DesiredWidth = -1 Then DesiredWidth = ctl.width
    
    border = 210
    minTabsPerRow = ctl.TabsPerRow
    tabsThisRow = 0
    rowsNeeded = 1
    
    ' Translate the captions, and compute how many tabs fit
    For i = 1 To ctl.Tabs
        Dim itemsize As size
        ctl.TabCaption(i - 1) = EW2W(ctl.TabCaption(i - 1))
        itemsize.cx = 0
        If ctl.TabVisible(i - 1) Then
            PixLocalizeGetStringSize frm.hdc, ctl.TabCaption(i - 1), itemsize
            itemsize.cx = itemsize.cx * Screen.TwipsPerPixelX + border
            widthSoFar = widthSoFar + itemsize.cx
            If widthSoFar > DesiredWidth Then
                rowsNeeded = rowsNeeded + 1
                If tabsThisRow < minTabsPerRow Then
                    minTabsPerRow = tabsThisRow
                End If
                tabsThisRow = 1
                widthSoFar = itemsize.cx
            Else
                tabsThisRow = tabsThisRow + 1
            End If
        End If
    Next i
    
    ' Move all controls if we need more rows
    If rowsNeeded > ctl.Rows Then
        Dim heightAdded As Long
        Dim widthAdded As Long
        heightAdded = (rowsNeeded - ctl.Rows) * ctl.TabHeight
        widthAdded = (rowsNeeded - ctl.Rows) * (10 * Screen.TwipsPerPixelX)
        ctl.TabsPerRow = minTabsPerRow
        frm.Height = frm.Height + heightAdded
        ctl.Height = ctl.Height + heightAdded
        frm.width = frm.width + widthAdded
        ctl.width = DesiredWidth + widthAdded
    
        Dim subctl As Control
        For Each subctl In frm.Controls
            Dim moveMe As Boolean
            
            If PixLocalizeIsChildOf(subctl, ctl) Then
                moveMe = True
            Else
                moveMe = False
            End If
            
            If moveMe Then
                subctl.Top = subctl.Top + heightAdded
            End If
        Next subctl
        
        For i = LBound(ctlsBelow) To UBound(ctlsBelow)
            ctlsBelow(i).Top = ctlsBelow(i).Top + heightAdded
        Next i
    End If

End Sub

Private Function isSpecialStyleHotkey(s As String) As Boolean
    Dim length As Long
    
    isSpecialStyleHotkey = False
    length = Len(s)
    If Mid(s, length, 1) = ")" Then
        If Mid(s, length - 2, 1) = "&" Then
            If Mid(s, length - 3, 1) = "(" Then
                isSpecialStyleHotkey = True
            End If
        End If
    End If
End Function

Public Function StripString(cap As String) As String
    Dim ret As String
    
    ret = cap
    ret = Trim(Replace(ret, ":", ""))
    If isSpecialStyleHotkey(ret) Then
        ret = Trim(Mid(ret, 1, Len(ret) - 4))
    Else
        ret = Replace(ret, "&", "")
    End If
    StripString = ret
End Function

Private Sub InitializeState(State As PixLocState)
    State.xScale = 1000
    State.yScale = 1000
    State.doRed = False
    State.doFont = False
    State.handleTabs = False
    State.numTabAnchors = 0
End Sub
