Attribute VB_Name = "modPictureFit"
Option Explicit

' Fits the currently selected picture inside the cell underneath the picture's
' upper-left corner. Store this module in PERSONAL.XLSB to use it in any workbook.
Private Const PICTURE_MARGIN_PERCENT As Double = 0.05

Public Sub FitSelectedPictureToTopLeftCell()
    Dim selectedPicture As Shape
    Dim targetRange As Range
    Dim validationMessage As String

    Set selectedPicture = GetSelectedPictureShape(validationMessage)
    If selectedPicture Is Nothing Then
        If Len(validationMessage) = 0 Then
            validationMessage = "Please select one pasted picture, then run this macro."
        End If

        MsgBox validationMessage, vbInformation, "Fit Picture to Cell"
        Exit Sub
    End If

    Set targetRange = GetTargetRangeFromPicture(selectedPicture)
    If targetRange Is Nothing Then
        MsgBox "The picture's target cell could not be identified.", vbInformation, "Fit Picture to Cell"
        Exit Sub
    End If

    If targetRange.Width <= 0 Or targetRange.Height <= 0 Then
        MsgBox "The target cell is hidden or too small to fit a picture.", vbInformation, "Fit Picture to Cell"
        Exit Sub
    End If

    On Error GoTo ResizeFailed
    FitPictureInsideRange selectedPicture, targetRange, PICTURE_MARGIN_PERCENT
    Exit Sub

ResizeFailed:
    MsgBox "The picture could not be resized. Check whether the worksheet is protected.", vbInformation, "Fit Picture to Cell"
End Sub

Private Function GetSelectedPictureShape(ByRef validationMessage As String) As Shape
    Dim selectedShape As Shape
    Dim selectedShapes As ShapeRange

    On Error Resume Next
    Set selectedShapes = Selection.ShapeRange
    On Error GoTo 0

    If selectedShapes Is Nothing Then
        validationMessage = "Please select one pasted picture, then run this macro."
        Exit Function
    End If

    If selectedShapes.Count <> 1 Then
        validationMessage = "Please select one picture only."
        Exit Function
    End If

    Set selectedShape = selectedShapes(1)

    If Not IsSupportedPictureShape(selectedShape) Then
        validationMessage = "The selected object is not a supported picture. Please select a pasted picture."
        Exit Function
    End If

    Set GetSelectedPictureShape = selectedShape
End Function

Private Function IsSupportedPictureShape(ByVal shp As Shape) As Boolean
    IsSupportedPictureShape = (shp.Type = msoPicture Or shp.Type = msoLinkedPicture)
End Function

Private Function GetTargetRangeFromPicture(ByVal pic As Shape) As Range
    Dim targetCell As Range

    Set targetCell = pic.TopLeftCell

    If targetCell.MergeCells Then
        Set GetTargetRangeFromPicture = targetCell.MergeArea
    Else
        Set GetTargetRangeFromPicture = targetCell
    End If
End Function

Private Sub FitPictureInsideRange(ByVal pic As Shape, ByVal targetRange As Range, ByVal marginPercent As Double)
    Dim availableWidth As Double
    Dim availableHeight As Double
    Dim widthScale As Double
    Dim heightScale As Double
    Dim scaleFactor As Double
    Dim newWidth As Double
    Dim newHeight As Double

    availableWidth = targetRange.Width * (1 - (2 * marginPercent))
    availableHeight = targetRange.Height * (1 - (2 * marginPercent))

    If availableWidth <= 0 Or availableHeight <= 0 Or pic.Width <= 0 Or pic.Height <= 0 Then
        Err.Raise vbObjectError + 513, "FitPictureInsideRange", "Invalid picture or target dimensions."
    End If

    widthScale = availableWidth / pic.Width
    heightScale = availableHeight / pic.Height
    scaleFactor = MinDouble(widthScale, heightScale)

    newWidth = pic.Width * scaleFactor
    newHeight = pic.Height * scaleFactor

    pic.LockAspectRatio = msoTrue
    pic.Width = newWidth
    pic.Height = newHeight
    pic.Left = targetRange.Left + ((targetRange.Width - newWidth) / 2)
    pic.Top = targetRange.Top + ((targetRange.Height - newHeight) / 2)
    pic.Placement = xlMove
End Sub

Private Function MinDouble(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue < secondValue Then
        MinDouble = firstValue
    Else
        MinDouble = secondValue
    End If
End Function
