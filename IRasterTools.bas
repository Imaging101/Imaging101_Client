Attribute VB_Name = "mod_IRasterTools"
' File:      IRasterTools.bas
' Created:   1998Nov26 by Rob Young
' Purpose:   To test the Spicer Edit Control's IRasterTools interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IRasterTools_BindToViewControl()
   'RobY Nov26/98
   Dim RasterTools As IRasterTools
   
   ' Set object variable for IRasterTools interface to Edit ctrl object
   Set RasterTools = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   'Bind the Spicer Edit Control to the Spicer View Control
   RasterTools.BindToViewControl MainMDIForm.ActiveForm.SpicerView1.object

   ' De-initialize the object variable
   Set RasterTools = Nothing
End Sub

Public Sub IRasterTools_DespeckleDialog()
   'RobY Nov26/98
   Dim RasterTools As IRasterTools
   
   ' Set object variable for IRasterTools interface to Edit ctrl object
   Set RasterTools = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   ' Display the despeckle dialog
   RasterTools.DespeckleDialog
   
   ' De-initialize the object variable
   Set RasterTools = Nothing
End Sub
Public Sub IRasterTools_RasterizeDialog()
   'RobY Nov26/98
   Dim RasterTools As IRasterTools
   
   ' Set object variable for IRasterTools interface to Edit ctrl object
   Set RasterTools = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   ' Display the rasterize dialog
   RasterTools.RasterizeDialog
   
   ' De-initialize the object variable
   Set RasterTools = Nothing
End Sub

Public Sub IRasterTools_ResizeDialog()
   'RobY Nov26/98
   Dim RasterTools As IRasterTools
   
   ' Set object variable for IRasterTools interface to Edit ctrl object
   Set RasterTools = MainMDIForm.ActiveForm.SpicerEdit1.object
   
   ' Display the resize dialog
   RasterTools.ResizeDialog
   
   ' De-initialize the object variable
   Set RasterTools = Nothing
End Sub


