Attribute VB_Name = "mod_ISpicerEdit"
' File:      ISpicerEdit.bas
' Created:   1998Nov5 by Rob Young
' Purpose:   To test the Spicer Edit Control's ISpicerEdit interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit

Public Sub ISpicerEdit_AboutBox()
   ' RobY Nov5/98
   MainMDIForm.ActiveForm.SpicerEdit1.AboutBox
End Sub

Public Sub ISpicerEdit_DataRotate180()
   ' RobY Nov26/98
   MainMDIForm.ActiveForm.SpicerEdit1.DataRotate180
End Sub

Public Sub ISpicerEdit_DataRotate90CW()
   ' RobY Nov27/98
   MainMDIForm.ActiveForm.SpicerEdit1.DataRotate90CW
End Sub

Public Sub ISpicerEdit_ResizeDialog()
   ' RobY Nov26/98
   MainMDIForm.ActiveForm.SpicerEdit1.ResizeDialog
End Sub

