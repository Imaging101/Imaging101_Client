Attribute VB_Name = "mod_IFileProperties"
' File:      IFileProperties.bas
' Created:   1998Nov25 by Mark Simpson
' Purpose:   To test the Spicer Document Control's IFileProperties interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IFileProperties_FileType()
   ' Modified by RobY Dec2/98 - Removed file format code for default interface
   Dim fileProperties As IFileProperties
   Dim fileFormat As FORMAT_TYPE

   ' Set object variable for IFileProperties interface to doc ctrl object
   Set fileProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
      
   ' Use the common dialog to allow user to select a file
   MainMDIForm.ActiveForm.CommonDialog1.Flags = cdlOFNHideReadOnly
   MainMDIForm.ActiveForm.CommonDialog1.DialogTitle = "ImageX - Select a File"
   MainMDIForm.ActiveForm.CommonDialog1.ShowOpen
   ' Get the selected file's file type
   fileFormat = fileProperties.FileType(MainMDIForm.ActiveForm.CommonDialog1.filename)
   ' Display result
   MsgBox "File: " + MainMDIForm.ActiveForm.CommonDialog1.filename + Chr(13) + "Returned file type:" + Str(fileFormat), vbInformation, "File Type"
   ' Initialize filename to null
   MainMDIForm.ActiveForm.CommonDialog1.filename = ""

   'De-initialize the object var
   Set fileProperties = Nothing
End Sub
   
Public Sub IFileProperties_FileTypeName()
   ' RobY Dec2/98
   Dim fileProperties As IFileProperties
   Dim strFileTypeName As String
   Dim iFileFormatNum As FORMAT_TYPE
   
   ' Set object variable for IFileProperties interface to doc ctrl object
   Set fileProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Prompt for file format number
   iFileFormatNum = InputBox("Please enter the file format number.", "File Type Number")
   strFileTypeName = fileProperties.FileTypeName(iFileFormatNum)
   ' display the format name
   MsgBox Str(iFileFormatNum) + " is " + strFileTypeName, vbInformation, "File Type Name"
   
   'De-initialize the object var
   Set fileProperties = Nothing
End Sub

Public Sub IFileProperties_FormatExtension()
   ' RobY Dec2/98
   Dim fileProperties As IFileProperties
   Dim strFormatExt As String
   Dim iFileFormatNum As Integer
   
   ' Set object variable for IFileProperties interface to doc ctrl object
   Set fileProperties = MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' Prompt for file format number
   iFileFormatNum = InputBox("Please enter the file format number.", "File Type Number")
   strFormatExt = fileProperties.FormatExtension(iFileFormatNum)
   ' display the format extension
   MsgBox "Format extension is " + strFormatExt + " for format type " + Str(iFileFormatNum) + ".", vbInformation, "Format Extension"
   
   'De-initialize the object var
   Set fileProperties = Nothing
End Sub
