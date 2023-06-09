Attribute VB_Name = "Spicer_Markup"
' File:      ISpicerMarkup.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Markup Control's ISpicerMarkup interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file.
'                          - ChildForm1 is no longer used access objects because it has been set to
'                             ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed
'                            to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub ISpicerMarkup_AboutBox()
   'RobY July30/98
   'Display the About Box for the Spicer Markup Control
   MainMDIForm.ActiveForm.SpicerMarkup1.AboutBox
End Sub

Public Sub ISpicerMarkup_BindToDocumentControl()
   'RobY July31/98
   'Bind the Spicer Markup Control to the Spicer Document Control
   MainMDIForm.ActiveForm.SpicerMarkup1.BindToDocumentControl MainMDIForm.ActiveForm.SpicerDoc1.object
End Sub

Public Sub ISpicerMarkup_BindToViewControl()
   'RobY July31/98
   'Bind the Spicer Markup Control to the Spicer Document Control
   MainMDIForm.ActiveForm.SpicerMarkup1.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
End Sub

