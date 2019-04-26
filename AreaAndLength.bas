Attribute VB_Name = "AreaAndLength"
 ' Written on free time during weekend 20-21 April 2019 upon requirement of Okazaki-san to have a simple app to measure area in the map
' 2019 Alvin Alexander
Option Explicit
Dim UnitLengthShape As Shape
Dim ShapeToMeasure As Shape
Dim MultiplierText As Shape
Sub Auto_Open()
    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
    Dim oButton2 As CommandBarButton
    Dim MyToolbar As String

    ' Give the toolbar a name
    MyToolbar = "Measure"

    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=MyToolbar, _
        Position:=msoBarFloating, Temporary:=True)
    If Err.Number <> 0 Then
          ' The toolbar's already there, so we have nothing to do
          Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Now add a button to the new toolbar
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    Set oButton2 = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties

    With oButton
         .DescriptionText = "Measure Area of a freeform close shape"
          'Tooltip text when mouse if placed over button
         .Caption = "Measure Area"
         'Text if Text in Icon is chosen
         .OnAction = "MeasureArea"
          'Runs the Sub Button1() code when clicked
         .Style = msoButtonIconAndCaption
          ' Button displays as icon, not text or both
         .FaceId = 174
          ' chooses icon #52 from the available Office icons
    End With
    
    With oButton2
         .DescriptionText = "Measure Length of a freeform open shape"
          'Tooltip text when mouse if placed over button
         .Caption = "Measure Length"
         'Text if Text in Icon is chosen
         .OnAction = "MeasureLength"
          'Runs the Sub Button1() code when clicked
         .Style = msoButtonIconAndCaption
          ' Button displays as icon, not text or both
         .FaceId = 130
          ' chooses icon #52 from the available Office icons
    End With

    ' Repeat the above for as many more buttons as you need to add
    ' Be sure to change the .OnAction property at least for each new button

    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Visible = True

NormalExit:
    Exit Sub   ' so it doesn't go on to run the errorhandler code

ErrorHandler:
     'Just in case there is an error
     MsgBox Err.Number & vbCrLf & Err.Description
     Resume NormalExit:
End Sub

Public Sub MeasureArea()
    Dim verticesArr() As Single
    Dim myArea As Double, finalArea As Double
    Set UnitLengthShape = Nothing
    Set ShapeToMeasure = Nothing
    Set MultiplierText = Nothing
    
    If Not InitialCheckOkay Then Exit Sub
    ' Check that the shape to measure shape is closed
    If Not ShapeIsClosed(ShapeToMeasure) Then MsgBox "Please select a closed shape.", vbCritical + vbOKOnly, "Selected shape is not closed": Exit Sub
    'Finally doing stuff
    verticesArr = ShapeToMeasure.vertices
    myArea = GetArea(verticesArr, UnitLengthShape.Width)
    finalArea = myArea * (GetMultiplier(MultiplierText) ^ 2)
    Call WriteAnswer(finalArea, ShapeToMeasure, " unit squared")
End Sub

Public Sub MeasureLength()
    Dim myLength As Double, finalLength As Double
    Dim verticesArr() As Single
    Set UnitLengthShape = Nothing
    Set ShapeToMeasure = Nothing
    Set MultiplierText = Nothing
    
    If Not InitialCheckOkay Then Exit Sub
    ' Check that the shape to measure shape is closed
    If ShapeIsClosed(ShapeToMeasure) Then MsgBox "Please select an open shape.", vbCritical + vbOKOnly, "Selected shape is closed": Exit Sub
    'Finally doing stuff
    verticesArr = ShapeToMeasure.vertices
    myLength = GetLength(verticesArr, UnitLengthShape.Width)
    finalLength = myLength * GetMultiplier(MultiplierText)
    Call WriteAnswer(finalLength, ShapeToMeasure, " units length")
End Sub

Private Sub WriteAnswer(value As Double, ShapeToMeasure As Shape, measureUnit As String)
    Dim oShp As Shape
    ' Write to screen the area
    With ActivePresentation.Slides(ActiveWindow.View.Slide.SlideIndex).Shapes
      Set oShp = .AddTextbox(msoTextOrientationHorizontal, ShapeToMeasure.Left, ShapeToMeasure.Top, 0, 0)
      With oShp
        .TextFrame2.TextRange.Text = value & " " & measureUnit
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .TextFrame2.WordWrap = msoFalse
      End With
    End With
End Sub

Private Function InitialCheckOkay() As Boolean
    Dim shapeCount As Integer, i As Integer
    InitialCheckOkay = True
    'Check if any shapes selected
    If Not ActiveWindow.Selection.Type = ppSelectionShapes Then MsgBox "Please select a unit-length horizontal line and a freeform shape to measure.", vbCritical + vbOKOnly, "No shape selected": InitialCheckOkay = False: Exit Function
    'Check if only and only if 2 or 3 shapes selected
    shapeCount = ActiveWindow.Selection.ShapeRange.Count
    If Not (shapeCount = 2 Or shapeCount = 3) Then MsgBox "Please select a line, a freeform, and a text box with multiplier value", vbCritical + vbOKOnly, "Wrong number of shapes selected": InitialCheckOkay = False: Exit Function
    
    'Set unit and to be measured shapes
    For i = 1 To shapeCount
        Select Case ActiveWindow.Selection.ShapeRange(i).Type
            Case msoLine
                Set UnitLengthShape = ActiveWindow.Selection.ShapeRange(i)
            Case msoFreeform
                Set ShapeToMeasure = ActiveWindow.Selection.ShapeRange(i)
            Case msoTextBox
                Set MultiplierText = ActiveWindow.Selection.ShapeRange(i)
        End Select
    Next
    ' Check at lease at least the 2 important shape available
    If UnitLengthShape Is Nothing Or ShapeToMeasure Is Nothing Then MsgBox "Not Correct combination of shapes selected.", vbCritical + vbOKOnly, "Wrong Shapes Selected": InitialCheckOkay = False: Exit Function
    ' Check unit length is perfectly horizontal
    If Not UnitLengthShape.Height = 0 Then MsgBox "No Perfectly Horizontal Line Selected", vbCritical + vbOKOnly, "Wrong Shapes Selected": InitialCheckOkay = False: Exit Function
    ' Check value of multiplier Text is a number
    If Not (MultiplierText Is Nothing) Then
        If Not IsNumeric(MultiplierText.TextFrame.TextRange.Text) Then MsgBox "No Number or not a number found in text box", vbCritical + vbOKOnly, "Wrong value in text box": InitialCheckOkay = False: Exit Function
    End If
End Function

' Check to see if the referenced shape has a closed path
Private Function ShapeIsClosed(oShp As Shape) As Boolean
    Dim FirstPoint() As Single, LastPoint() As Single
    Dim myVertices() As Single
    
    ' Save all of the vertices from the shape to an array
    myVertices = oShp.vertices
    
    ' Check the the first node's coordinates are the same as the last node e.g. closed path
    FirstPoint = oShp.Nodes(1).Points
    LastPoint = oShp.Nodes(oShp.Nodes.Count).Points
    If FirstPoint(1, 1) = LastPoint(1, 1) And FirstPoint(1, 2) = LastPoint(1, 2) Then ShapeIsClosed = True
End Function
 
' Algorithm to calculate the area of an irregular polygon using its vertices
Private Function GetArea(ByRef myVertices() As Single, ByRef oneUnit As Double) As Double
    Dim counter As Integer
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, mySum1 As Double, mySum2 As Double
    Dim unitsq As Double
    For counter = LBound(myVertices) To UBound(myVertices) - 1
        x1 = myVertices(counter, 1)
        y1 = myVertices(counter, 2)
        x2 = myVertices(counter + 1, 1)
        y2 = myVertices(counter + 1, 2)
        mySum1 = mySum1 + (x1 * y2)
        mySum2 = mySum2 + (x2 * y1)
    Next
    unitsq = oneUnit ^ 2
    GetArea = Round((Abs(mySum1 - mySum2) / 2) / unitsq, 4)
End Function

Private Function GetLength(ByRef myVertices() As Single, ByRef oneUnit As Double) As Double
    Dim counter As Integer
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, mySum1 As Double
    For counter = LBound(myVertices) To UBound(myVertices) - 1
        x1 = myVertices(counter, 1)
        y1 = myVertices(counter, 2)
        x2 = myVertices(counter + 1, 1)
        y2 = myVertices(counter + 1, 2)
        mySum1 = mySum1 + (((x1 - x2) ^ 2) + ((y1 - y2) ^ 2)) ^ 0.5
    Next
    GetLength = Round(mySum1 / oneUnit, 4)
End Function

Private Function GetMultiplier(TextBoxShape As Shape) As Double
    Dim multiplier As Variant
    If TextBoxShape Is Nothing Then GetMultiplier = 1: Exit Function
    GetMultiplier = TextBoxShape.TextFrame.TextRange.Text
End Function
' Measuring Area is Adopted and modified from
' VBA Macro to calculate the area of the selected irregular polygon shape
'----------------------------------------------------------------------------------
' Copyright (c) 2014 YOUpresent Ltd.
' Source code is provide under Creative Commons Attribution License
' This means you must give credit for our original creation in the following form:
' "Includes code created by YOUpresent Ltd. (YOUpresent.co.uk)"
' Commons Deed @ http://creativecommons.org/licenses/by/3.0/
' License Legal @ http://creativecommons.org/licenses/by/3.0/legalcode
'----------------------------------------------------------------------------------
' Purpose : Calculates the are of the selected shape and optionally adds a square
'           of the same area or a user-selected percentage of it.
'
' Author : Jamie Garroch
' Date : 08SEP2014
' Website : http://youpresent.co.uk and http://www.gmark.co
'----------------------------------------------------------------------------------
