Attribute VB_Name = "mod_ISpicerView"
' File:      ISpicerView.bas
' Created:   1998Apr28 by Rob Wood, Mark Simpson, and Rob Young
' Purpose:   To test the Spicer View Control's ISpicerView interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.

Option Explicit

Public Sub ISpicerView_AboutBox()
   ' RobY May5/98
   ' Display the SpicerView Control About Box
   MainMDIForm.ActiveForm.SpicerView1.AboutBox
End Sub

Public Sub ISpicerView_ActivePageID()
   ' RobY May5/98
   ' Display the Active Page ID
   MsgBox ("Active Page ID: " + Str(MainMDIForm.ActiveForm.SpicerView1.ActivePageId) + " (Should NOT be zero)")
End Sub

Public Sub ISpicerView_BindToDocumentControl()
   ' RobY May5/98
   ' Bind doc control to the view control
   MainMDIForm.ActiveForm.SpicerView1.BindToDocumentControl (MainMDIForm.ActiveForm.SpicerDoc1.object)
End Sub

Public Sub ISpicerView_Enabled()
   ' RobY May5/98
   ' Displays the value for the Enabled status and sets it
   MsgBox ("Enabled Status: " + Str(MainMDIForm.ActiveForm.SpicerView1.Enabled))
End Sub

Public Sub ISpicerView_hWnd()
   ' RobY May5/98
   ' Displays the window handle(hWnd)
   MsgBox ("Window Handle(hWnd): " + Str(MainMDIForm.ActiveForm.SpicerView1.hWnd))
End Sub

' <EOF ISpicerView.bas>

