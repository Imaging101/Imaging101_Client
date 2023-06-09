Attribute VB_Name = "mod_ISimplifyView"
' File:      ISimplifyView.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer View Control's ISimplifyView interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit
Dim MenuName As Menu

Public Sub ISimplifyView_Hairlines()
   ' RobW May 12/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lRootID As Long
   Dim sTemp As String
   
   ' Get root ID for current doc
   lRootID = MainMDIForm.ActiveForm.SpicerDoc1.RootID
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, using document as object
   bOriginal = SimplifyView.Hairlines(lRootID)
   ' Set hairlines to opposite setting / toggle the setting
   SimplifyView.Hairlines(lRootID) = Not bOriginal
   sTemp = "Set hairlines to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.Hairlines(lRootID)
   sTemp = sTemp + "Get hairlines returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_Invert()
   ' RobW May 12/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.Invert(lPageID)
   ' Set invert to opposite setting / toggle the setting
   SimplifyView.Invert(lPageID) = Not bOriginal
   sTemp = "Set invert to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.Invert(lPageID)
   sTemp = sTemp + "Get invert returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_MaximizeAnnotations()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Call the MaximizeAnnotations method
   SimplifyView.MaximizeAnnotations
   
   ' De-initialize object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_MaximizeAnnotationsAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim SimplifyView As ISimplifyView
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for ISimplifyView interface to doc ctrl object
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_ISimplifyView_Array(2)
   
   ' Get the value of availability
   iAvailable = SimplifyView.MaximizeAnnotationsAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_MinimizeAnnotations()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Call the MinimizeAnnotations method
   SimplifyView.MinimizeAnnotations
   
   ' De-initialize object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_MinimizeAnnotationsAvailability()
   'RobY July1/98
   ' Modified by RobY Dec11/98 - changed global function to change menu status of ImageX according to what is returned from availablity
   Dim SimplifyView As ISimplifyView
   Dim iAvailable As COMMAND_AVAILABILITY
   
   ' Set object variable for ISimplifyView interface to doc ctrl object
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Set the menu name to change
   Set MenuName = MainMDIForm.ActiveForm.mnu_ViewCtrl_ISimplifyView_Array(3)
   
   ' Get the value of availability
   iAvailable = SimplifyView.MinimizeAnnotationsAvailability
   ' Change the menu status through function
   Availability iAvailable, MenuName
   
   ' De-initialize object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_Mirror()
   ' RobW May 12/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.Mirror(lPageID)
   ' Set mirror to opposite setting / toggle the setting
   SimplifyView.Mirror(lPageID) = Not bOriginal
   sTemp = "Set mirror to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.Mirror(lPageID)
   sTemp = sTemp + "Get mirror returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_Monochrome()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.Monochrome(lPageID)
   ' Set monochrome to opposite setting / toggle the setting
   SimplifyView.Monochrome(lPageID) = Not bOriginal
   sTemp = "Set monochrome to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.Monochrome(lPageID)
   sTemp = sTemp + "Get monochrome returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_Negative()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.Negative(lPageID)
   ' Set negative to opposite setting / toggle the setting
   SimplifyView.Negative(lPageID) = Not bOriginal
   sTemp = "Set negative to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.Negative(lPageID)
   sTemp = sTemp + "Get negative returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_RowAndColumnDisplay()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.RowAndColumnDisplay(lPageID)
   ' Set RowAndColumn to opposite setting / toggle the setting
   SimplifyView.RowAndColumnDisplay(lPageID) = Not bOriginal
   sTemp = "Set RowAndColumnDisplay to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.RowAndColumnDisplay(lPageID)
   sTemp = sTemp + "Get RowAndColumnDisplay returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_Sample()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.Sample(lPageID)
   ' Set sample to opposite setting / toggle the setting
   SimplifyView.Sample(lPageID) = Not bOriginal
   sTemp = "Set sample to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.Sample(lPageID)
   sTemp = sTemp + "Get sample returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_ShowAnnotations()
   ' RobW May 13/98
   ' Modified by RobY July10/98
   Dim SimplifyView As ISimplifyView
   Dim strOriginal As String
   Dim strNewSetting As String
   Dim lPageID As Long
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   strOriginal = SimplifyView.ShowAnnotations(lPageID)
   ' Set to opposite of original
   If strOriginal = False Then
      SimplifyView.ShowAnnotations(lPageID) = True
   Else
      SimplifyView.ShowAnnotations(lPageID) = False
   End If
   ' Get new setting for current page
   strNewSetting = SimplifyView.ShowAnnotations(lPageID)
   If strOriginal = strNewSetting Then
      MsgBox "Originally set to: " + strOriginal + Chr(13) + "Now set to:" + strNewSetting, vbCritical, "Failed"
   Else
      MsgBox "Originally set to: " + strOriginal + Chr(13) + "Now set to:" + strNewSetting, vbInformation, "Passed"
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_ShowEraseOutlines()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.ShowEraseOutlines(lPageID)
   ' Set ShowEraseOutlines to opposite setting / toggle the setting
   SimplifyView.ShowEraseOutlines(lPageID) = Not bOriginal
   sTemp = "Set ShowEraseOutlines to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.ShowEraseOutlines(lPageID)
   sTemp = sTemp + "Get ShowEraseOutlines returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_ShowHotspots()
   ' RobW May 13/98
   ' Modified by RobY July10/98
   Dim SimplifyView As ISimplifyView
   Dim strOriginal As String
   Dim strNewSetting As String
   Dim lPageID As Long
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   strOriginal = SimplifyView.ShowHotspots(lPageID)
   ' Set to opposite of original
   If strOriginal = False Then
      SimplifyView.ShowHotspots(lPageID) = True
   Else
      SimplifyView.ShowHotspots(lPageID) = False
   End If
   ' Get new setting for current page
   strNewSetting = SimplifyView.ShowHotspots(lPageID)
   If strOriginal = strNewSetting Then
      MsgBox "Originally set to: " + strOriginal + Chr(13) + "Now set to:" + strNewSetting, vbCritical, "Failed"
   Else
      MsgBox "Originally set to: " + strOriginal + Chr(13) + "Now set to:" + strNewSetting, vbInformation, "Passed"
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_ShowPasteOutlines()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.ShowPasteOutlines(lPageID)
   ' Set ShowPasteOutlines to opposite setting / toggle the setting
   SimplifyView.ShowPasteOutlines(lPageID) = Not bOriginal
   sTemp = "Set ShowPasteOutlines to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.ShowPasteOutlines(lPageID)
   sTemp = sTemp + "Get ShowPasteOutlines returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_StampVisible()
   ' RobW May 13/98
   ' Modified by RobY June23/98
   Dim SimplifyView As ISimplifyView
   Dim iOriginal As Integer
   Dim iNewSetting As Integer
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   'Get current setting
   iOriginal = SimplifyView.StampVisible(lPageID)
   If iOriginal = 0 Then
      sTemp = "StampVisible was set to OFF." + Chr(13)
   Else
      sTemp = "StampVisible was set to ON." + Chr(13)
   End If
   ' Toggle between on and off
   SimplifyView.StampVisible(lPageID) = IN_TOGGLE
   ' Get new setting
   iNewSetting = SimplifyView.StampVisible(lPageID)
   If iNewSetting = 0 Then
      sTemp = sTemp + "StampVisible now set to OFF."
   Else
      sTemp = sTemp + "StampVisible now set to ON."
   End If
   ' Check if setting was switched / if it worked
   If iNewSetting <> iOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

Public Sub ISimplifyView_Wireframes()
   ' RobW May 13/98
   Dim SimplifyView As ISimplifyView
   Dim bOriginal As Boolean
   Dim bNewSetting As Boolean
   Dim lPageID As Long
   Dim sTemp As String
   
   ' Get Page ID for current page, current doc
   lPageID = MainMDIForm.ActiveForm.SpicerView1.ActivePageId
   ' Set object variable for ISimplifyView interface to the view control
   Set SimplifyView = MainMDIForm.ActiveForm.SpicerView1.object
   ' Get current setting, for current page
   bOriginal = SimplifyView.Wireframes(lPageID)
   ' Set wireframes to opposite setting / toggle the setting
   SimplifyView.Wireframes(lPageID) = Not bOriginal
   sTemp = "Set WireFrames to " + Str(Not bOriginal) + "." + Chr(13)
   ' Get new setting
   bNewSetting = SimplifyView.Wireframes(lPageID)
   sTemp = sTemp + "Get WireFrames returned" + Str(bNewSetting) + "."
   ' Check if setting was switched / if it worked
   If bNewSetting = Not bOriginal Then
      ' Worked
      MsgBox sTemp, vbInformation
   Else
      ' Didn't work
      MsgBox sTemp, vbCritical
   End If
   
   ' De-initialize the object var
   Set SimplifyView = Nothing
End Sub

' <EOF ISimplifyView.bas>
