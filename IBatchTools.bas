Attribute VB_Name = "mod_IBatchTools"
' File:      IBatchTools.bas
' Created:   1998July30 by Rob Young
' Purpose:   To test the Spicer Reference Control's IBatchTools interface.
' Revisions: RobY Jan4/99  - Added functionality to open more than 1 file. ChildForm1 is no longer used access objects because it has been set to ChildForm() which is an array of ChildForm1's forms.
'                          - Anywhere ChildForm1 was used has been changed to ChildForm(iIndex) or MainMDIForm.ActiveForm.
'   $log$
'   $nokeywords$
'
' Copyright (C) 1998 SPICER Corporation, All rights reserved.
Option Explicit

Public Sub IBatchTools_BindToDocumentControl()
   'RobY July31/98
   Dim BatchTools As IBatchTools
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   'Bind the Spicer Markup Control to the Spicer Document Control
   BatchTools.BindToDocumentControl MainMDIForm.ActiveForm.SpicerDoc1.object
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_ExtractText()
   'RobY July31/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim strFilename As String
   Dim lCount As Long
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
      
   lLayerID = InputBox("Please enter the layer id for the text objects you want to extract.", "Layer ID")
   strFilename = InputBox("Please enter the path and filename where you want to save the new file.", "Filename")
   'Extract the text
   BatchTools.ExtractText lLayerID, strFilename, lCount
   MsgBox "Number of text objects extracted is" + Str(lCount), vbInformation, "Text Objects Extracted"
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_NumberOfTextReplaced()
   ' RobY Dec11/98
   ' Modified by RobY Dec16/98 - Moved from IVectorProperties module
   Dim BatchTools As IBatchTools
   Dim lNumTextReplaced As Long
   
   ' Set object variable for IBatchTools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   ' Get the number of text replaced by the TextReplace command
   lNumTextReplaced = BatchTools.NumberOfTextReplaced
   ' Display number with msgbox
   MsgBox "The number of text replaced by the TextReplace method is " + Str(lNumTextReplaced) + ".", vbInformation, "NumberOfTextReplaced"
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceAnnotation()
   'RobY Aug10/98
   'Modified by RobY Nov20/98 Removed centered parameter
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX As Double
   Dim dY As Double
   Dim strString As String
   Dim lColor As Long
   Dim iTypeFace As TEXT_TYPEFACE
   Dim strFontName As String
   Dim dWidth As Double
   Dim dHeight As Double
   Dim iUnitType As UNIT_TYPE
   Dim dRotation As Double
   Dim iJustification As TEXT_JUSTIFICATION
   Dim iMirror As MIRROR_STATE
   Dim bIconized As Boolean
   Dim dOrientAngle As Double
   Dim dShearAngle As Double
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 1
   lParentID = 0
   dX = 0.395
   dY = 1.695
   strString = "ImageX Test for PlaceAnnotation"
   lColor = 255
   iTypeFace = IN_TYPEFACE_NORMAL
   strFontName = "Times New Roman"
   dWidth = 0.1
   dHeight = 0.2
   iUnitType = IN_UNITS_INCH
   dRotation = 0
   iJustification = IN_JUST_LEFT + IN_JUST_TOP
   iMirror = IN_MIRROR_NONE
   bIconized = True
   dOrientAngle = 0
   dShearAngle = 0
   
   'Place an annotation on the active layer
   BatchTools.PlaceAnnotation lLayerID, iVectorState, lObjectID, lParentID, dX, dY, strString, _
      lColor, iTypeFace, strFontName, dWidth, dHeight, iUnitType, dRotation, _
      iJustification, iMirror, bIconized, dOrientAngle, dShearAngle
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceArc()
   'RobY Aug19/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dStartAngle As Double
   Dim dEndAngle As Double
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iLineFillOp As LINE_FILL_OPERATION
   Dim iCapStyle As LINE_CAP_STYLE
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 0
   lParentID = 0
   dStartAngle = 180
   dEndAngle = 360
   dX1 = 2.4
   dY1 = -1.1
   dX2 = 1.36
   dY2 = 1.33
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   iLineStyle = IN_LINE_SOLID
   iLineFillOp = IN_ROP_OPAQUE
   iCapStyle = IN_CAP_ROUND
   
   'Place an arc on the active layer
   BatchTools.PlaceArc lLayerID, iVectorState, lObjectID, lParentID, dStartAngle, dEndAngle, dX1, dY1, dX2, dY2, dThickness, _
      iThicknessUnits, lLineColor, iLineStyle, iLineFillOp, iCapStyle
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceBox()
   'RobY Aug19/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iJointStyle As LINE_JOIN_STYLE
   Dim iFillStyle As FILL_STYLE
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 5
   lParentID = 0
   dX1 = 2.3
   dY1 = 1.2
   dX2 = 2.6
   dY2 = 1.3
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iJointStyle = IN_JOIN_ROUND
   iFillStyle = IN_FILL_TRANSLUCENT
   
   'Place a box on the active layer
   BatchTools.PlaceBox lLayerID, iVectorState, lObjectID, lParentID, dX1, dY1, dX2, dY2, dThickness, _
      iThicknessUnits, lLineColor, lFillColor, iLineStyle, iJointStyle, iFillStyle
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceCircle()
   'RobY Aug11/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dRadius As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iFillStyle As FILL_STYLE
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 6
   lParentID = 0
   dX1 = 2.3
   dY1 = 1.2
   dRadius = 3
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iFillStyle = IN_FILL_TRANSLUCENT
   
   'Place a circle on the active layer
   BatchTools.PlaceCircle lLayerID, iVectorState, lObjectID, lParentID, dX1, dY1, dRadius, dThickness, _
      iThicknessUnits, lLineColor, lFillColor, iLineStyle, iFillStyle
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceEllipse()
   'RobY Aug19/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iFillStyle As FILL_STYLE
   Dim dRotation As Double
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 0
   lParentID = 0
   dX1 = 25.8
   dY1 = 8.2
   dX2 = 15.2
   dY2 = 3.3
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iFillStyle = IN_FILL_TRANSPARENT
   dRotation = 0
   
   'Place a ellipse on the active layer
   BatchTools.PlaceEllipse lLayerID, iVectorState, lObjectID, lParentID, dX1, dY1, dX2, dY2, dThickness, _
      iThicknessUnits, lLineColor, lFillColor, iLineStyle, iFillStyle, dRotation
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceGroup()
   'RobY Aug19/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim bUnbindable As Boolean
   Dim bMoveable As Boolean
   Dim bScaleable As Boolean
   Dim strGroupID As String
   Dim ltest As Long
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 14
   lParentID = 0
   bUnbindable = True
   bMoveable = True
   bScaleable = True
   strGroupID = 0
   
   'Place a group on the active layer
   BatchTools.PlaceGroup lLayerID, iVectorState, lObjectID, lParentID, bUnbindable, bMoveable, _
            bScaleable, strGroupID

   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub
   
Public Sub IBatchTools_PlaceLine()
   'RobY Aug11/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim dX2 As Double
   Dim dY2 As Double
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iLineFillOp As LINE_FILL_OPERATION
   Dim iCapStyle As LINE_CAP_STYLE
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 0
   lParentID = 0
   dX1 = 1.1
   dY1 = 1.5
   dX2 = 1.7
   dY2 = 1.5
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lColor = 255
   iLineStyle = IN_LINE_SOLID
   iLineFillOp = IN_ROP_OPAQUE
   iCapStyle = IN_CAP_ROUND
   
   'Place a line on the active layer
   BatchTools.PlaceLine lLayerID, iVectorState, lObjectID, lParentID, dX1, dY1, dX2, dY2, dThickness, _
      iThicknessUnits, lColor, iLineStyle, iLineFillOp, iCapStyle
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlacePolygon()
   'RobY Aug19/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim iAdditionalPoints As Integer
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim lFillColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iJointStyle As LINE_JOIN_STYLE
   Dim iFillStyle As FILL_STYLE
   Dim iPolyFillMode As POLY_FILL_MODE
   Dim iCurveStyle As CURVE_STYLE
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 0
   lParentID = 0
   dX1 = 1
   dY1 = 1.1
   iAdditionalPoints = 4
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   lFillColor = 255
   iLineStyle = IN_LINE_SOLID
   iJointStyle = IN_JOIN_ROUND
   iFillStyle = IN_FILL_TRANSPARENT
   iPolyFillMode = IN_POLYFILL_WINDING
   iCurveStyle = IN_CURVE_NONE
   
   'Place a polygon on the active layer
   BatchTools.PlacePolygon lLayerID, iVectorState, lObjectID, lParentID, dX1, dY1, iAdditionalPoints, dThickness, _
      iThicknessUnits, lLineColor, lFillColor, iLineStyle, iJointStyle, iFillStyle, iPolyFillMode, iCurveStyle
   BatchTools.AddPoint 1.6, 1.1
   BatchTools.AddPoint 1.6, 1.3
   BatchTools.AddPoint 1, 1.3
   BatchTools.AddPoint 4, 2.3

   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlacePolyline()
   'RobY Aug19/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim iAdditionalPoints As Integer
   Dim dThickness As Double
   Dim iThicknessUnits As UNIT_TYPE
   Dim lLineColor As Long
   Dim iLineStyle As LINE_STYLE
   Dim iLineFillOp As LINE_FILL_OPERATION
   Dim iCapStyle As LINE_CAP_STYLE
   Dim iJointStyle As LINE_JOIN_STYLE
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 0
   lParentID = 0
   iAdditionalPoints = 10
   dX1 = 1.045
   dY1 = 1.975
   dThickness = 0.005
   iThicknessUnits = IN_UNITS_INCH
   lLineColor = 255
   iLineStyle = IN_LINE_SOLID
   iLineFillOp = IN_ROP_OPAQUE
   iCapStyle = IN_CAP_ROUND
   iJointStyle = IN_JOIN_ROUND
   
   'Place a polyline on the active layer
   BatchTools.PlacePolyline lLayerID, iVectorState, lObjectID, lParentID, dX1, dY1, iAdditionalPoints, dThickness, _
      iThicknessUnits, lLineColor, iLineStyle, iLineFillOp, iCapStyle, iJointStyle
   BatchTools.AddPoint 1.05, 3.975
   BatchTools.AddPoint 1.05, 3.98
   BatchTools.AddPoint 1.055, 3.98
   BatchTools.AddPoint 1.055, 5.98
   BatchTools.AddPoint 1.055, 3.98
   BatchTools.AddPoint 1.055, 3.985
   BatchTools.AddPoint 1.065, 3.995
   BatchTools.AddPoint 1.065, 2
   BatchTools.AddPoint 1.07, 2
   BatchTools.AddPoint 1.075, 2.01

   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceSymbol()
   'RobY Aug20/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim lX As Double
   Dim lY As Double
   Dim strFilename As String
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   lX = 50000
   lY = 50000
   strFilename = InputBox("Please enter the path and filename of the symbol to place.", "Filename")
   
   'Place a symbol on the active layer
   BatchTools.PlaceSymbol lLayerID, IN_UNITS_PROPORTIONAL, lX, lY, strFilename
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_PlaceText()
   'RobY Aug19/98
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim iVectorState As VECTOR_STATE
   Dim lObjectID As Long
   Dim lParentID As Long
   Dim dX1 As Double
   Dim dY1 As Double
   Dim strText As String
   Dim lTextColor As Long
   Dim iTypeFace As TEXT_TYPEFACE
   Dim strFontName As String
   Dim dWidth As Double
   Dim dHeight As Double
   Dim iUnitType As UNIT_TYPE
   Dim dRotation As Double
   Dim iJustification As TEXT_JUSTIFICATION
   Dim iMirror As MIRROR_STATE
   Dim dOrientAngle As Double
   Dim dShearAngle As Double
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object

   lLayerID = MainMDIForm.ActiveForm.SpicerMarkup1.ActiveLayer
   iVectorState = IN_OBJFLAG_VISIBLE + IN_OBJFLAG_NEW
   lObjectID = 0
   lParentID = 0
   dX1 = 0.1
   dY1 = 1.4
   strText = "PlaceText worked!!!!"
   lTextColor = 255
   iTypeFace = IN_TYPEFACE_NORMAL
   strFontName = "Times New Roman"
   dWidth = 0.2
   dHeight = 0.4
   iUnitType = IN_UNITS_INCH
   dRotation = 0
   iJustification = IN_JUST_LEFT + IN_JUST_TOP
   iMirror = IN_MIRROR_NONE
   dOrientAngle = 0
   dShearAngle = 0
   
   'Place text on the active layer
   BatchTools.PlaceText lLayerID, iVectorState, lObjectID, lParentID, dX1, dY1, strText, _
         lTextColor, iTypeFace, strFontName, dWidth, dHeight, _
         iUnitType, dRotation, iJustification, iMirror, dOrientAngle, dShearAngle
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_TextReplace()
   ' RobY Oct28/98
   ' Modified by RobY Dec2/98 - Setup so user can specify values
   ' Modified by Roby Dec10/98 - Rewrote code to reflect parameter changes
   ' Modified by RobY Dec16/98 - Moved from IVectorProperties module
   Dim BatchTools As IBatchTools
   Dim lOccurrence As Long
   Dim iLayerNum As Integer
   Dim lLayerID As Long
   Dim strSearchString As String
   Dim strReplaceString As String
   
   ' Set object variable for IBatchTools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
   
   ' Prompt user for layer # to search
   iLayerNum = InputBox("Please enter the layer number of the objects to search and replace the text for.", "Layer Number")
   ' Get the layer id of the specified layer
   lLayerID = MainMDIForm.ActiveForm.SpicerDoc1.LayerID(MainMDIForm.ActiveForm.SpicerView1.ActivePageId, iLayerNum)
   ' Prompt for string to replace
   strSearchString = InputBox("Please enter the string that you want to replace.", "Current String")
   ' Prompt for string to use as replacement
   strReplaceString = InputBox("Please enter the new string to use.", "New String")
   ' Prompt for number of occurrences
   lOccurrence = InputBox("Please enter the number of occurrences to search for and then replace.", "Occurrences")
   ' Search and replace the the matching text
   BatchTools.TextReplace lLayerID, strSearchString, strReplaceString, lOccurrence
   MsgBox "Any matching string have been replaced with the new string.", vbInformation, "TextReplace"
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub

Public Sub IBatchTools_TextSearch()
   Dim BatchTools As IBatchTools
   Dim lLayerID As Long
   Dim strStringToFind As String
   Dim lOccurrence As Long
   Dim lMatchCount As Long
   Dim iObjectType As TOOL_TYPE
   Dim strToolType As String
   Dim lX1 As Long
   Dim lY1 As Long
   Dim lX2 As Long
   Dim lY2 As Long
   Dim strMatchString As String
   Dim strReturn As String
   
   ' Set object variable for IBatchtools interface to Markup ctrl object
   Set BatchTools = MainMDIForm.ActiveForm.SpicerMarkup1.object
      
   strReturn = Chr(10) + Chr(13)
   
   lLayerID = InputBox("Please enter the layer id for the edit layer you want to search.", "Layer ID")
   strStringToFind = InputBox("Please enter the string to search for.", "Text Search")
   lOccurrence = MsgBox("Do you want to search for all occurrences?", vbInformation + vbYesNo, "Find All Matches")
   If lOccurrence = vbYes Then
      lOccurrence = 0
   Else
      lOccurrence = InputBox("Please enter the number of occurences to search for.", "Amount to Search")
   End If
   'Search for the string
   BatchTools.TextSearch lLayerID, strStringToFind, lOccurrence, lMatchCount, iObjectType, lX1, lY1, lX2, lY2, strMatchString
   Select Case iObjectType
      Case IN_TOOL_ANNOTATION
         strToolType = "ANNOTATION"
      Case IN_TOOL_TEXT
         strToolType = "Text"
      Case Else
         strToolType = Error
   End Select
   MsgBox "Match Count =" + Str(lMatchCount) + strReturn + _
          "Tool Type =" + strToolType + strReturn + _
          "x1 = " + Str(lX1) + strReturn + "y1 = " + Str(lY1) + strReturn + "x2 =" + Str(lX2) + _
          strReturn + "y2 =" + Str(lY2) + strReturn + _
          "String Match =" + strMatchString, vbInformation, "String Matches Found"
   
   ' De-initialize the object variable
   Set BatchTools = Nothing
End Sub



