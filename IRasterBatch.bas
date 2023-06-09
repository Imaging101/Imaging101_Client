Attribute VB_Name = "mod_IRasterBatch"
' File:      IRasterBatch.bas
' Created:   1998Nov26 by Rob Young
' Purpose:   To test the Spicer Edit Control's IRasterBatch interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IRasterBatch_BindToDocumentControl()
   'RobY Nov26/98
   Dim RasterBatch As IRasterBatch
   
   ' Set object variable for IRasterBatch interface to Edit ctrl object
   Set RasterBatch = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   'Bind the Spicer Edit Control to the Spicer View Control
   RasterBatch.BindToDocumentControl MainMDIForm.ActiveForm.SpicerDoc1.object

   ' De-initialize the object variable
   Set RasterBatch = Nothing
End Sub

Public Sub IRasterBatch_Deskew()
   'RobY Nov26/98
   Dim RasterBatch As IRasterBatch
   Dim dAngle As Double
   
   ' Set object variable for IRasterBatch interface to Edit ctrl object
   Set RasterBatch = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   ' Prompt for deskew angle
   dAngle = InputBox("Please enter the angle that you want to deskew the image at.", "Deskew Angle")
   ' Deskew at the specified angle
   RasterBatch.Deskew dAngle
   
   ' De-initialize the object variable
   Set RasterBatch = Nothing
End Sub

Public Sub IRasterBatch_ProcessOperations()
   'MarkS Nov 27, 1998
   Dim RasterBatch As IRasterBatch
   
   ' Set object variable for IRasterBatch interface to Edit ctrl object
   Set RasterBatch = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   ' Cannot test ProcessOperations yet, being worked on at the C-API level by Brad
   'RasterBatch.ProcessOperations
   
   
   ' De-initialize the object variable
   Set RasterBatch = Nothing
End Sub
Public Sub IRasterBatch_Rasterize()
   'RobY Nov26/98
   Dim RasterBatch As IRasterBatch
   Dim iMergeType As RASTERIZE_TYPE
   Dim iXResolution As Integer
   Dim iYResolution As Integer
   Dim bColor As Boolean
   Dim bDither As Boolean
   
   ' Set object variable for IRasterBatch interface to Edit ctrl object
   Set RasterBatch = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   iMergeType = IN_RASTERIZE_DOCUMENT
   iXResolution = 400
   iYResolution = 400
   bColor = False
   bDither = False
   
   ' Rasterize the document
   RasterBatch.Rasterize iMergeType, iXResolution, iYResolution, bColor, bDither
   MsgBox "The document has been rasterized with these settings: " + Chr(13) + _
            "MergeType = " + Str(iMergeType) + Chr(13) + "X & Y Resolution = " + Str(iXResolution) + _
            Chr(13) + "Color = " + Str(bColor) + Chr(13) + "Dither = " + Str(bDither), vbInformation, "Rasterize"
            
   ' De-initialize the object variable
   Set RasterBatch = Nothing
End Sub

