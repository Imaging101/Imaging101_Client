Attribute VB_Name = "mod_IINIFilePreferences"
' File:      IINIFilePreferences.bas
' Created:   1998Nov30 by Rob Young
' Purpose:   To test the Spicer Configuration Control's IINIFilePreferences interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IINIFilePreferences_VectorLoadRasterMapping()
   'RobY Nov30/98
   Dim INIFilePref As IINIFilePreferences
   Dim bCurSetting As Boolean
   Dim bNewSetting As Boolean
   
   ' Set object variable for IINIFilePreferences interface to Configuration ctrl object
   Set INIFilePref = MainMDIForm.ActiveForm.SpicerConfiguration1.object
   
   ' Get the current setting
   bCurSetting = INIFilePref.VectorLoadRasterMapping
   ' Change to opposite of current setting
   If bCurSetting = False Then
      INIFilePref.VectorLoadRasterMapping = True
   Else
      INIFilePref.VectorLoadRasterMapping = False
   End If
   ' Get new setting
   bNewSetting = INIFilePref.VectorLoadRasterMapping
   If bNewSetting = bCurSetting Then
      MsgBox "Failed to change to new setting." + Chr(13) + "Old Setting =" + Str(bCurSetting) + _
               Chr(13) + "New Setting =" + Str(bNewSetting), vbCritical, "Failed"
   Else
      MsgBox "Successfully changed to new setting." + Chr(13) + "Old Setting =" + Str(bCurSetting) + _
               Chr(13) + "New Setting =" + Str(bNewSetting), vbInformation, "Success"
   End If
   
   ' De-initialize the object variable
   Set INIFilePref = Nothing
End Sub
