Attribute VB_Name = "mod_ISpicerReference"
' File:      ISpicerReference.bas
' Created:   1998July13 by Rob Young
' Purpose:   To test the Spicer Reference Control's ISpicerReference interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub ISpicerReference_AboutBox()
   'RobY July/98
   'Display the About Box for the Spicer Reference control
   frmReference.SpicerReference1.AboutBox
End Sub

Public Sub ISpicerReference_BindToViewControl()
   'RobY July/98
   'Bind the Spicer Reference control to the Spicer View control
   frmReference.SpicerReference1.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
End Sub

Public Sub ISpicerReference_HighlightBoxColor()
   'RobY Sept21/98
   'Modified by RobY Dec14/98 - Allow box colour to be changed
   Dim lCurColour As Long
   Dim iResponse As Integer
   Dim lNewColour As Long
   Dim lGetNewColour As Long
   
   ' Get current box colour
   lCurColour = frmReference.SpicerReference1.HighlightBoxColor
   ' Display current colour and prompt if user wants to change
   iResponse = MsgBox("Current 24-bit Colour Value: " + Str(lCurColour) + Chr(13) + _
                  "Do you want to change the box colour?", vbInformation + vbYesNo, "Highlight Box Colour")
   If iResponse = vbYes Then
      ' Prompt for new colour
      lNewColour = InputBox("Please enter new 24-bit colour value.", "New Highlight Box Colour")
      ' Change to new colour
      frmReference.SpicerReference1.HighlightBoxColor = lNewColour
      ' Get new colour value
      lGetNewColour = frmReference.SpicerReference1.HighlightBoxColor
      If lNewColour = lGetNewColour Then
         MsgBox "Highlight box colour has been changed.", vbInformation, "Highlight Box Colour Changed"
      Else
         MsgBox "Highlight box colour was not changed to what was set." + Chr(13) + "Set New Colour: " + _
                  Str(lNewColour) + Chr(13) + "Colour Returned: " + Str(lGetNewColour), vbCritical, "ERROR"
      End If
   Else
      MsgBox "Highlight box colour was not changed.", vbInformation, "No Change"
   End If
End Sub
