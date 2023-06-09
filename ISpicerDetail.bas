Attribute VB_Name = "mod_ISpicerDetail"
' File:      ISpicerReference.bas
' Created:   1998July13 by Rob Young
' Purpose:   To test the Spicer Detail Control's ISpicerDetail interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub ISpicerDetail_AboutBox()
   'RobY July/98
   'Display the About Box for the Spicer Detail control
   frmDetail.SpicerDetail1.AboutBox
End Sub

Public Sub ISpicerDetail_BindToViewControl()
   'RobY July/98
   'Bind the Spicer Detail control to the Spicer View control
   frmDetail.SpicerDetail1.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object
End Sub

