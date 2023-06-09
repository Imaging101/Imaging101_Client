Attribute VB_Name = "Spicer_UserTools"
' File:      IUserTools.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Markup Control's IUserTools interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file.
'                          - ChildForm1 is no longer used access objects because it has been set to
'                             ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed
'                             to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.


Option Explicit
Dim MenuName As Menu

Public Sub IUserTools_ActiveLayer()
   'RobY Aug11/98
   Dim UserTools As IUserTools
   Dim lSetLayerID As Long
   Dim lGetLayerID As Long
   Dim strReturn As String
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   strReturn = Chr(13) + Chr(10)
   'Prompt for layer to activate
   lSetLayerID = InputBox("Please enter the layer id that you want to make active.", "Layer Activate")
   'Set new active layer
   UserTools.ActiveLayer = lSetLayerID
   'Get new active layer id
   lGetLayerID = UserTools.ActiveLayer
   If lGetLayerID = lSetLayerID Then
      MsgBox "Layer id set to " + Str(lGetLayerID), vbInformation, "Active Layer ID"
   Else
      MsgBox "Get does not match what was set." + strReturn + "Set = " + Str(lSetLayerID) + strReturn + "Get = " + Str(lGetLayerID), vbCritical, "Active Layer ID "
   End If
   
   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_ActiveLayerDialog()
   'RobY Aug11/98
   Dim UserTools As IUserTools
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Display the dialog
   UserTools.ActiveLayerDialog
   
   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_ActiveLayerDialogAvailability()
   'RobY Dec11/98
   Dim UserTools As IUserTools
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IUserTools_Array(1)
   
   ' Get the value of availability
   iAvailable = UserTools.ActiveLayerDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_ActiveTool(Index As Integer)
   'RobY Aug11/98
   ' Modified by RobY Dec29/98 - Changed so user can select tool through
   '  Menu List
   Dim UserTools As IUserTools
   Dim iSelectedTool As Integer
   Dim strReturn As String
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
         
   If Index = 0 Then
      UserTools.ActiveTool = 1   ' Cut
   ElseIf Index = 1 Then
      UserTools.ActiveTool = 2   ' Copy
   ElseIf Index = 2 Then
      UserTools.ActiveTool = 3   ' Paste
   ElseIf Index = 4 Then
      UserTools.ActiveTool = 17  ' Rubout
   ElseIf Index = 5 Then
      UserTools.ActiveTool = 18  ' Erase Area
   ElseIf Index = 7 Then
      UserTools.ActiveTool = 4   ' Line
   ElseIf Index = 8 Then
      UserTools.ActiveTool = 8   ' Arrow
   ElseIf Index = 9 Then
      UserTools.ActiveTool = 9   ' Sketch
   ElseIf Index = 10 Then
      UserTools.ActiveTool = 10  ' Polyline
   ElseIf Index = 11 Then
      UserTools.ActiveTool = 29  ' Arc
   ElseIf Index = 12 Then
      UserTools.ActiveTool = 5   ' Box
   ElseIf Index = 13 Then
      UserTools.ActiveTool = 6   ' Circle
   ElseIf Index = 14 Then
      UserTools.ActiveTool = 7   ' Ellipse
   ElseIf Index = 15 Then
      UserTools.ActiveTool = 11  ' Polygon
   ElseIf Index = 16 Then
      UserTools.ActiveTool = 12  ' Text
   ElseIf Index = 17 Then
      UserTools.ActiveTool = 13  ' Annotation
   ElseIf Index = 18 Then
      UserTools.ActiveTool = 31  ' Highlighter
   ElseIf Index = 19 Then
      UserTools.ActiveTool = 32  ' Highlight Area
   ElseIf Index = 20 Then
      UserTools.ActiveTool = 14  ' Dimension
   ElseIf Index = 21 Then
      UserTools.ActiveTool = 15  ' Symbol
   ElseIf Index = 22 Then
      UserTools.ActiveTool = 16  ' Hotspot
   ElseIf Index = 24 Then
      UserTools.ActiveTool = 19  ' Select
   ElseIf Index = 25 Then
      UserTools.ActiveTool = 20  ' Select All
   ElseIf Index = 26 Then
      UserTools.ActiveTool = 21  'Deselect All
   ElseIf Index = 27 Then
      UserTools.ActiveTool = 24  ' Delete Selected Objects
   ElseIf Index = 28 Then
      UserTools.ActiveTool = 26  ' Bind
   ElseIf Index = 29 Then
      UserTools.ActiveTool = 27  ' Unbind
   ElseIf Index = 30 Then
      UserTools.ActiveTool = 22  ' Move/Resize
   ElseIf Index = 31 Then
      UserTools.ActiveTool = 23  ' Rotate
   ElseIf Index = 32 Then
      UserTools.ActiveTool = 25  ' Save As Symbol
   ElseIf Index = 33 Then
      UserTools.ActiveTool = 28  ' Change Text
   ElseIf Index = 34 Then
      UserTools.ActiveTool = 33  ' Change Hotspot
   End If
   
   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_ActiveToolAvailability()
   'RobY Dec11/98
   'Modified by RobY Dec29/98 - Changed so that the availability for each
   '  tool is returned and the menu is updated
   Dim UserTools As IUserTools
   Dim Available As COMMAND_AVAILABILITY
   Dim iCount As Integer
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
  
   ' Set menu status in ActiveToolAvailability menu
   For iCount = 0 To 34
      If (iCount <> 3) And (iCount <> 6) And (iCount <> 23) Then
         ' Set the menu name to change
         Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IUserTools_ActiveTool_Array(iCount)
         ' Get availability for each tool
         If iCount = 0 Then
            Available = UserTools.ToolAvailability(1)    ' Cut
         ElseIf iCount = 1 Then
            Available = UserTools.ToolAvailability(2)    ' Copy
         ElseIf iCount = 2 Then
            Available = UserTools.ToolAvailability(3)    ' Paste
         ElseIf iCount = 4 Then
            Available = UserTools.ToolAvailability(17)   ' Rubout
         ElseIf iCount = 5 Then
            Available = UserTools.ToolAvailability(18)   ' Erase Area
         ElseIf iCount = 7 Then
            Available = UserTools.ToolAvailability(4)    ' Line
         ElseIf iCount = 8 Then
            Available = UserTools.ToolAvailability(8)    ' Arrow
         ElseIf iCount = 9 Then
            Available = UserTools.ToolAvailability(9)    ' Sketch
         ElseIf iCount = 10 Then
            Available = UserTools.ToolAvailability(10)  ' Polyline
         ElseIf iCount = 11 Then
            Available = UserTools.ToolAvailability(29)  ' Arc
         ElseIf iCount = 12 Then
            Available = UserTools.ToolAvailability(5)   ' Box
         ElseIf iCount = 13 Then
            Available = UserTools.ToolAvailability(6)   ' Circle
         ElseIf iCount = 14 Then
            Available = UserTools.ToolAvailability(7)   ' Ellipse
         ElseIf iCount = 15 Then
            Available = UserTools.ToolAvailability(11)  ' Polygon
         ElseIf iCount = 16 Then
            Available = UserTools.ToolAvailability(12)  ' Text
         ElseIf iCount = 17 Then
            Available = UserTools.ToolAvailability(13)  ' Annotation
         ElseIf iCount = 18 Then
            Available = UserTools.ToolAvailability(31)  ' Highlighter
         ElseIf iCount = 19 Then
            Available = UserTools.ToolAvailability(32)  ' Highlight Area
         ElseIf iCount = 20 Then
            Available = UserTools.ToolAvailability(14)  ' Dimension
         ElseIf iCount = 21 Then
            Available = UserTools.ToolAvailability(15)  ' Symbol
         ElseIf iCount = 22 Then
            Available = UserTools.ToolAvailability(16)  ' Hotspot
         ElseIf iCount = 24 Then
            Available = UserTools.ToolAvailability(19)  ' Select
         ElseIf iCount = 25 Then
            Available = UserTools.ToolAvailability(20)  ' Select All
         ElseIf iCount = 26 Then
            Available = UserTools.ToolAvailability(21)  'Deselect All
         ElseIf iCount = 27 Then
            Available = UserTools.ToolAvailability(24)  ' Delete Selected Objects
         ElseIf iCount = 28 Then
            Available = UserTools.ToolAvailability(26)  ' Bind
         ElseIf iCount = 29 Then
            Available = UserTools.ToolAvailability(27)  ' Unbind
         ElseIf iCount = 30 Then
            Available = UserTools.ToolAvailability(22)  ' Move/Resize
         ElseIf iCount = 31 Then
            Available = UserTools.ToolAvailability(23)  ' Rotate
         ElseIf iCount = 32 Then
            Available = UserTools.ToolAvailability(25)  ' Save As Symbol
         ElseIf iCount = 33 Then
            Available = UserTools.ToolAvailability(28)  ' Change Text
         ElseIf iCount = 34 Then
             Available = UserTools.ToolAvailability(33)  ' Change Hotspot
         End If
         ' Set the menu status through Availability procedure in mod_Global.bas module
         Availability Available, MenuName
      End If
   Next iCount
   
   ' De-initialize object var
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_BindToViewControl()
   'RobY Aug11/98
    Dim UserTools As IUserTools
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   'Bind the Spicer Markup Control to the Spicer View Control
   UserTools.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
   
   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_MoveLayer()
   'RobY Aug11/98
   Dim UserTools As IUserTools
   Dim lLayerID As Long
   Dim lRelLayerID As Long
   Dim lX1 As Long
   Dim lY1 As Long
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   lX1 = 25000
   lY1 = 25000
   
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, 2)
   lRelLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, 1)
   'Move layer 2 to be 1/4 of layer 1 in both directions
   UserTools.MoveLayer lLayerID, lRelLayerID, lX1, lY1
   
   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_MoveLayerDialog()
   'RobY Aug11/98
    Dim UserTools As IUserTools
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   'Display the move layer dialog
   UserTools.MoveLayerDialog
   
   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_MoveLayerDialogAvailability()
   'RobY Dec11/98
   Dim UserTools As IUserTools
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IUserTools_Array(7)
   
   ' Get the value of availability
   iAvailable = UserTools.MoveLayerDialogAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_SnapToGrid()
   ' RobY Nov20/98
   ' Modified by RobY Dec11/98 - Moved from IVectorProperties module to
   '   IUserTools module. Change code to reflect changes.
   Dim UserTools As IUserTools
   Dim iSetting As TOGGLE_BOOL
   Dim strSetting As String

   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Toggle the setting
   UserTools.SnapToGrid = IN_TOGGLE
   ' Get the current setting
   iSetting = UserTools.SnapToGrid
   ' Convert to string
   Select Case iSetting
      Case IN_ON
         strSetting = "ON"
      Case IN_OFF
         strSetting = "OFF"
   End Select
   MsgBox "SnapToGrid turned " + strSetting, vbInformation, "SnapToGrid"

   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_SnapToRightAngles()
   ' RobY Nov20/98
   ' Modified by RobY Dec11/98 - Moved from IVectorProperties module to
   '   IUserTools module. Change code to reflect changes.
   Dim UserTools As IUserTools
   Dim iSetting As TOGGLE_BOOL
   Dim strSetting As String
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   'Toggle the setting
   UserTools.SnapToRightAngles = IN_TOGGLE
   ' Get the current setting
   iSetting = UserTools.SnapToRightAngles
   ' Convert to string
   Select Case iSetting
      Case IN_ON
         strSetting = "ON"
      Case IN_OFF
         strSetting = "OFF"
   End Select
   MsgBox "SnapToRightAngles turned " + strSetting, vbInformation, "SnapToRightAngles"
   
   ' De-initialize the object variable
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_SnapToGridAvailability()
   'RobY Dec11/98
   Dim UserTools As IUserTools
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IUserTools_Array(9)
   
   ' Get the value of availability
   iAvailable = UserTools.SnapToGridAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set UserTools = Nothing
End Sub

Public Sub IUserTools_SnapToRightAnglesAvailability()
   'RobY Dec11/98
   Dim UserTools As IUserTools
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IUserTools interface to Markup ctrl object
   Set UserTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_Markup_IUserTools_Array(10)
   
   ' Get the value of availability
   iAvailable = UserTools.SnapToRightAnglesAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set UserTools = Nothing
End Sub




