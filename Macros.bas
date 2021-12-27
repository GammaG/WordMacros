Attribute VB_Name = "NewMacros"
Sub DeleteSmallPictures()
Dim iShp As InlineShape
    For Each iShp In ActiveDocument.InlineShapes
        With iShp
          
                iShp.Delete
        
        End With
    Next iShp
End Sub
Sub FigureInfo()
    Dim iShapeCount As Integer
    Dim iILShapeCount As Integer
    Dim DocThis As Document
    Dim J As Integer
    Dim sTemp As String

    Set DocThis = ActiveDocument
    Documents.Add

    iShapeCount = DocThis.Shapes.Count
    If iShapeCount > 0 Then
        Selection.TypeText Text:="Regular Shapes"
        Selection.TypeParagraph
    End If
    For J = 1 To iShapeCount
        Selection.TypeText Text:=DocThis.Shapes(J).Name
        Selection.TypeParagraph
        sTemp = "     Height (points): "
        sTemp = sTemp & DocThis.Shapes(J).Height
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        sTemp = "     Width (points): "
        sTemp = sTemp & DocThis.Shapes(J).Width
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        sTemp = "     Height (pixels): "
        sTemp = sTemp & PointsToPixels(DocThis.Shapes(J).Height, True)
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        sTemp = "     Width (pixels): "
        sTemp = sTemp & PointsToPixels(DocThis.Shapes(J).Width, False)
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        Selection.TypeParagraph
    Next J

    iILShapeCount = DocThis.InlineShapes.Count
    If iILShapeCount > 0 Then
        Selection.TypeText Text:="Inline Shapes"
        Selection.TypeParagraph
    End If
    For J = 1 To iILShapeCount
        Selection.TypeText Text:="Shape " & J
        Selection.TypeParagraph
        sTemp = "     Height (points): "
        sTemp = sTemp & DocThis.InlineShapes(J).Height
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        sTemp = "     Width (points): "
        sTemp = sTemp & DocThis.InlineShapes(J).Width
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        sTemp = "     Height (pixels): "
        sTemp = sTemp & PointsToPixels(DocThis.InlineShapes(J).Height, True)
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        sTemp = "     Width (pixels): "
        sTemp = sTemp & PointsToPixels(DocThis.InlineShapes(J).Width, False)
        Selection.TypeText Text:=sTemp
        Selection.TypeParagraph
        Selection.TypeParagraph
    Next J
End Sub
Sub RemoveTextBox1()
    Dim oShpNrm As Shape
    Dim oInlineShpNrm As InlineShape
    Dim sText As String
    On Error GoTo ErrHandler
    For Each oShpNrm In ActiveDocument.Shapes
    sText = "no Word Art"
sText = oShpNrm.TextEffect.Text
If sText <> "no Word Art" Then
oShpNrm.Select
Selection.Text = sText
End If
Next
For Each oInlineShpNrm In ActiveDocument.InlineShapes
sText = "no Word Art"
sText = oInlineShpNrm.TextEffect.Text
If sText <> "no Word Art" Then
oInlineShpNrm.Select
Selection.Text = sText
End If
Next
ErrHandler:
Err.Clear
Resume Next
 
Reply

End Sub


Sub RemoveWordArt()

Dim oShpNrm As Shape
Dim oInlineShpNrm As InlineShape
Dim sText As String
On Error GoTo ErrHandler
For Each oShpNrm In ActiveDocument.Shapes
sText = "no Word Art"
sText = oShpNrm.TextEffect.Text
If sText <> "no Word Art" Then
oShpNrm.Select
Selection.Text = sText
End If
Next
For Each oInlineShpNrm In ActiveDocument.InlineShapes
sText = "no Word Art"
sText = oInlineShpNrm.TextEffect.Text
If sText <> "no Word Art" Then
oInlineShpNrm.Select
Selection.Text = sText
End If
Next
ErrHandler:
Err.Clear
Resume Next

End Sub
Sub RemoveHeaderAndFooter()

    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter

    For Each oSec In ActiveDocument.Sections
        For Each oHead In oSec.Headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
End Sub

