Attribute VB_Name = "mod_IMiscellaneous"
' File:      IMiscellaneous.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer View Control's IMiscellaneous interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub IMiscellaneous_ActivateAllHotspots()
   'Mark Simpson May 19/1998
   'Modified by RobY June22/98
   Dim lLayerID As Long
   Dim Misc As IMiscellaneous
   Dim iAnswer As Integer
   
   iAnswer = MsgBox("Do you know the layer ID where the hotspot(s) that you want activated are?", vbYesNo)
   If iAnswer = vbYes Then
     'Set object variable for IMiscellaneous interface to view control
      Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
      
      lLayerID = Val(InputBox("Please enter the layer ID that contains the hotspot(s) that you want activated?"))
      'Activate All Hotspots on the specified layer
      Misc.ActivateAllHotspots lLayerID
   
      'De-initialize the object variable
      Set Misc = Nothing
   Else
      MsgBox "To get the layer ID, use the LayerID method in the IDocContents interface."
   End If
End Sub

Public Sub IMiscellaneous_ActivateHotspot()
   'Mark Simpson May 19/1998
   'Modified by RobY Dec2/98 - Removed error message and wrote code for command
   Dim Misc As IMiscellaneous
   Dim lLayerID As Long
   Dim strHotspotID As String
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   lLayerID = Val(InputBox("Please enter the layer ID that contains the hotspot that you want activate.", "LayerID"))
   strHotspotID = InputBox("Please enter the hotspot id of the hotspot to activate.", "HotspotID")
   'Activate All Hotspots on the specified layer
   Misc.ActivateHotspot lLayerID, strHotspotID

   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_ActivateMeasureTool()
   'Mark Simpson May 14/98
   Dim Misc As IMiscellaneous
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   'Activate the measure tool
   Misc.ActivateMeasureTool
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_ActivateMeasureToolAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim Miscellaneous As IMiscellaneous
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IMiscellaneous interface to doc ctrl object
   Set Miscellaneous = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IMiscellaneous_Array(2)
   
   ' Get the value of availability
   iAvailable = Miscellaneous.ActivateMeasureToolAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set Miscellaneous = Nothing
End Sub

Public Sub IMiscellaneous_CopyDocumentToOleObject()
   'Mark Simpson May 14/98
   Dim Misc As IMiscellaneous
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   'Copy Document to OLE Object - copies the current page
   Misc.CopyDocumentToOleObject
   
   'Send message to try and paste as an OLE object
   MsgBox "Current page was just copied to OLE object" + Chr(13) + Chr(13) + "Now try and paste the page into another program.", vbInformation
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_CopyDocumentToOleObjectAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim Miscellaneous As IMiscellaneous
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IMiscellaneous interface to doc ctrl object
   Set Miscellaneous = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IMiscellaneous_Array(3)
   
   ' Get the value of availability
   iAvailable = Miscellaneous.CopyDocumentToOleObjectAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set Miscellaneous = Nothing
End Sub

Public Sub IMiscellaneous_DisplayGrid()
   'Mark Simpson May 14/98
   Dim Misc As IMiscellaneous
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   'Display the grid
   Misc.DisplayGrid
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_DisplayGridAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim Miscellaneous As IMiscellaneous
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IMiscellaneous interface to doc ctrl object
   Set Miscellaneous = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IMiscellaneous_Array(4)
   
   ' Get the value of availability
   iAvailable = Miscellaneous.DisplayGridAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set Miscellaneous = Nothing
End Sub

Public Sub IMiscellaneous_DocumentScrollBars()
   'Roby Nov27/98
   Dim Misc As IMiscellaneous
   Dim iScrollBars As SCROLL_BAR
   Dim strDocScrollBars As String
   Dim iResponse As Integer
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Prompt for scroll bar setting
   iScrollBars = InputBox("Please enter a scroll bar setting." + Chr(13) + "IN_FLAG_SCROLLBAR_NONE = 0" + Chr(13) + _
                                 "IN_FLAG_SCROLLBAR_HORZ = 1" + Chr(13) + "IN_FLAG_SCROLLBAR_VERT = 2" + Chr(13) + "IN_FLAG_SCROLLBAR_BOTH = 3", "Document Scroll Bars")
   Misc.DocumentScrollBars = iScrollBars
   strDocScrollBars = ScrollBars(iScrollBars)
   iResponse = MsgBox("DocumentScrollBars set to " + strDocScrollBars + "." + Chr(13) + "Did the scroll bar(s) display correctly?", vbYesNo, "Document Scroll Bars")
   If iResponse = vbYes Then
      MsgBox "Scroll bars were successfully changed.", vbInformation, "Success"
   Else
      MsgBox "Changing scroll bars failed.", vbInformation, "Failed"
   End If
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub
Public Sub IMiscellaneous_DocwinID()
   'Roby July1/98
   Dim Misc As IMiscellaneous
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   'Display the docwin ID
   MsgBox "Docwin ID = " + Str(Misc.DocwinID), vbInformation, "Docwin ID"
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_GetObjectExtents()
   'Roby Nov27/98
   Dim Misc As IMiscellaneous
   Dim lObjectID As Long
   Dim dReturnX1 As Double
   Dim dReturnY1 As Double
   Dim dReturnX2 As Double
   Dim dReturnY2 As Double
   Dim iUnitType As Integer
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
      
   ' Prompt for object id
   lObjectID = InputBox("Please enter the page or layer id to get the extents of." + Chr(13) + _
                        "(Returns extents of what is displayed)", "Object ID")
   ' Prompt for unit type
   iUnitType = InputBox("Please enter the unit type you want the extents returned in." + Chr(13) + _
                        Chr(13) + "INCH = 1" + Chr(13) + "CM = 2" + Chr(13) + "FT = 4" + Chr(13) + "MM = 5" + _
                        Chr(13) + "M = 6", "Unit Type")
   ' Get the extents
   Misc.GetObjectExtents lObjectID, iUnitType, dReturnX1, dReturnY1, dReturnX2, dReturnY2
   MsgBox "Object Extents:" + Chr(13) + "x1=" + Str(dReturnX1) + Chr(13) + "y1=" + Str(dReturnY1) + Chr(13) + _
            "x2=" + Str(dReturnX2) + Chr(13) + "y2=" + Str(dReturnY2), vbInformation, "GetObjectExtents"
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_MoveGrid()
   'Mark Simpson May 14/98
   Dim Misc As IMiscellaneous
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   'Move the grid
   Misc.MoveGrid
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_MoveGridAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim Miscellaneous As IMiscellaneous
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for IMiscellaneous interface to doc ctrl object
   Set Miscellaneous = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_IMiscellaneous_Array(8)
   
   ' Get the value of availability
   iAvailable = Miscellaneous.MoveGridAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set Miscellaneous = Nothing
End Sub

Public Sub IMiscellaneous_PageScrollBars()
   'Roby Nov26/98
   Dim Misc As IMiscellaneous
   Dim iScrollBars As SCROLL_BAR
   Dim strPageScrollBars As String
   Dim iResponse As Integer
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Prompt for scroll bar setting
   iScrollBars = InputBox("Please enter a scroll bar setting." + Chr(13) + "IN_FLAG_SCROLLBAR_NONE = 0" + Chr(13) + _
                                 "IN_FLAG_SCROLLBAR_HORZ = 1" + Chr(13) + "IN_FLAG_SCROLLBAR_VERT = 2" + Chr(13) + "IN_FLAG_SCROLLBAR_BOTH = 3", "Page Scroll Bars")
   Misc.PageScrollBars(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) = iScrollBars
   strPageScrollBars = ScrollBars(iScrollBars)
   iResponse = MsgBox("PageScrollBars set to " + strPageScrollBars + "." + Chr(13) + "Did the scroll bar(s) display correctly?", vbYesNo, "Page Scroll Bars")
   If iResponse = vbYes Then
      MsgBox "Scroll bars were successfully changed.", vbInformation, "Success"
   Else
      MsgBox "Changing scroll bars failed.", vbInformation, "Failed"
   End If
   
   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_OCRRegion()
   'Mark Simpson May 19/1998
   'Modified by RobY July10/98
   Dim Misc As IMiscellaneous
   Dim strFilename As String
   Dim x1, y1, x2, y2 As Long
   Dim iFormatID As FORMAT_TYPE
   
   ' Set object variable for IMiscellaneous interface to doc ctrl object
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   strFilename = InputBox("Please enter name and location of where to save file.", "Save")
   iFormatID = InputBox("Please enter value of format id.", "Format ID")
   x1 = InputBox("Please enter value of x1 in proportional units.", "X1 Value")
   y1 = InputBox("Please enter value of y1 in proportional units.", "Y1 Value")
   x2 = InputBox("Please enter value of x2 in proportional units.", "X2 Value")
   y2 = InputBox("Please enter value of y2 in proportional units.", "Y2 Value")
   Misc.OCRRegion strFilename, iFormatID, IN_UNITS_PROPORTIONAL, x1, y1, x2, y2
   
   ' De-initialize object var
   Set Misc = Nothing
End Sub

Public Sub IMiscellaneous_ViewPreferencesDialog()
   'Mark Simpson May 14/98
   ' Modified by RobY Dec 16/98 - Rewrote code to include availability and to allow user to specify scope
   Dim Misc As IMiscellaneous
   Dim Scope As PROPERTY_SCOPE
   Dim Available As COMMAND_AVAILABILITY
   
   'Set object variable for IMiscellaneous interface to view control
   Set Misc = MainMDIForm.ActiveForm.SpicerView1.object
   
   ' Prompt for scope to use
   Scope = InputBox("Please enter the scope that you want to use. Check help file for values.", "Scope")
   ' Get the value of availability
   Available = Misc.ViewPreferencesDialogAvailability(Scope)
   ' Do the case statement that corresponds to the availability returned
   Select Case Available
      Case IN_UI_ENABLED
         MsgBox "Command is ENABLED. Click OK to display dialog.", vbInformation, "Availability"
         ' Display the dialog
         Misc.ViewPreferencesDialog Scope
      Case IN_UI_GREYED
         MsgBox "Command is GREYED.", vbInformation, "Availability"
      Case IN_UI_REMOVED
         MsgBox "Command is REMOVED.", vbInformation, "Availability"
      Case IN_UI_CHECKED
         MsgBox "Command is CHECKED.", vbInformation, "Availability"
      Case IN_UI_ENABLED_CHECKED
         MsgBox "Command is ENABLED and CHECKED. Click OK to display dialog.", vbInformation, "Availability"
         ' Display the dialog
         Misc.ViewPreferencesDialog Scope
      Case Else
         MsgBox "Cannot open dialog. Check to make that the scope you specified is valid.", vbExclamation, "Availability"
   End Select

   'De-initialize the object variable
   Set Misc = Nothing
End Sub

Public Function ScrollBars(iScrollBars As SCROLL_BAR) As String
   ' RobY Nov26/98
   Dim strScrollBars As String
   
   Select Case iScrollBars
      Case IN_FLAG_SCROLLBAR_NONE
         strScrollBars = "NONE"
      Case IN_FLAG_SCROLLBAR_HORZ
         strScrollBars = "HORIZONTAL"
      Case IN_FLAG_SCROLLBAR_VERT
         strScrollBars = "VERTICAL"
      Case IN_FLAG_SCROLLBAR_BOTH
         strScrollBars = "BOTH"
   End Select
   ScrollBars = strScrollBars
End Function
