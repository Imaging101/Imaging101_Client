Attribute VB_Name = "mod_ICFGConfiguration"
' File:      ICFGConfiguration.bas
' Created:   1998Nov30 by Rob Young
' Purpose:   To test the Spicer Edit Control's IRasterTools interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub ICFGConfiguration_AllowMouseZoomInWindow()
   'RobY Nov30/98
   Dim CFGConfig As ICFGConfiguration
   Dim bCurSetting As Boolean
   Dim bNewSetting As Boolean
   
   ' Set object variable for ICFGConfiguration interface to Configuration ctrl object
   Set CFGConfig = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   ' Get current setting
   bCurSetting = CFGConfig.AllowMouseZoomInWindow
   ' Set to opposite of current setting
   If bCurSetting = False Then
      CFGConfig.AllowMouseZoomInWindow = True
   Else
      CFGConfig.AllowMouseZoomInWindow = False
   End If
   ' Get new setting
   bNewSetting = CFGConfig.AllowMouseZoomInWindow
   
   ' De-initialize the object variable
   Set CFGConfig = Nothing
End Sub

Public Sub ICFGConfiguration_AllowRightContextMenuInWindow()
   'RobY Nov30/98
   Dim CFGConfig As ICFGConfiguration
   Dim bCurSetting As Boolean
   Dim bNewSetting As Boolean
   
   ' Set object variable for ICFGConfiguration interface to Configuration ctrl object
   Set CFGConfig = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   ' Get current setting
   bCurSetting = CFGConfig.AllowRightContextMenuInWindow
   ' Set to opposite of current setting
   If bCurSetting = False Then
      CFGConfig.AllowRightContextMenuInWindow = True
   Else
      CFGConfig.AllowRightContextMenuInWindow = False
   End If
   ' Get new setting
   bNewSetting = CFGConfig.AllowRightContextMenuInWindow
   If bCurSetting = bNewSetting Then
      MsgBox "Failed to change to new setting." + Chr(13) + "Old Setting: " + Str(bCurSetting) + _
            "New Setting: " + Str(bNewSetting), vbCritical, "Failed"
   Else
      MsgBox "Changed to new setting." + Chr(13) + "Old Setting: " + Str(bCurSetting) + _
            Chr(13) + "New Setting: " + Str(bNewSetting), vbInformation, "Success"
   End If
   
   ' De-initialize the object variable
   Set CFGConfig = Nothing
End Sub
