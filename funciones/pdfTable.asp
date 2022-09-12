<%
Class Table ' only needed in ASP

Private mDoc
Private mRect
Private mRectTop
Private mRectLeft
Private mRectWidth
Private mRectHeight
Private mRowTop
Private mRowBottom
Private mHeights()
Private mXPos
Private mYPos
Private mColumns
Private mWidths()
Private mObjects
Private mTruncated
Public RowHeightMin
Public RowHeightMax
Public Padding

' Focus on the document and assign the relevant number of columns
Public Sub Focus(inDoc, inColumns)
  Set mDoc = inDoc
  SetRect mDoc.Rect
  SetColumns inColumns
End Sub

' Add a new page, reset the table rect and move to the first row
Public Sub NewPage()
  mDoc.Page = mDoc.AddPage
  SetRect mRect
  NextRow
End Sub

' Assign a new table rectangle and reset the current table position
Public Sub SetRect(inRect)
  mDoc.Rect = inRect
  mRect = mDoc.Rect
  mRectTop = mDoc.Rect.Top
  mRectLeft = mDoc.Rect.Left
  mRectWidth = mDoc.Rect.Width
  mRectHeight = mDoc.Rect.Height
  mRowTop = mDoc.Rect.Top
  mRowBottom = mDoc.Rect.Top
  ReDim mHeights(0)
  mYPos = -1
  mXPos = -1
End Sub

' Change the number of columns in the table
Public Sub SetColumns(inNum)
  Dim i
  If inNum > 0 Then
    ReDim Preserve mWidths(inNum - 1)
    For i = (mColumns - 1) To (inNum - 1)
      If i >= 0 Then mWidths(i) = 1
    Next
    mColumns = inNum
  End If
End Sub

' Get the current row - a zero based index
Public Property Get Row()
  Row = mYPos
End Property

' Get the current column - a zero based index
Public Property Get Column()
  Column = mXPos
End Property

' Find out if the last row we added was truncated
Public Property Get RowTruncated()
  RowTruncated = mTruncated
End Property

' Change a column width
Public Property Let Width(i, inWidth)
  If i >= 0 And i <= UBound(mWidths) Then mWidths(i) = inWidth
End Property

' Move to the next column in the current row
Public Sub NextCell()
  mXPos = mXPos + 1
  If mXPos >= mColumns Then mXPos = mColumns - 1
  If mXPos < 0 Then mXPos = 0
  SelectCurrentCell mXPos
End Sub

' Move to the next row - return false if the next row would not fit
Public Function NextRow()
  Dim theBottom
  mRowTop = mRowTop - GetRowHeight(mYPos)
  mDoc.Rect = mRect
  mRowBottom = mRowTop
  mDoc.Rect.Top = mRowTop
  If RowHeightMax > 0 Then
    theBottom = mRowTop - RowHeightMax
    If mDoc.Rect.Bottom < theBottom Then mDoc.Rect.Bottom = theBottom
  End If
  NextRow = (mDoc.Rect.Height > RowHeightMin + (2 * Padding))
  If NextRow Then
    mYPos = mYPos + 1
    mXPos = -1
    mObjects = ""
    'mTruncated = False
  End If
End Function

' Add text to the currently selected area
Public Function AddText(inText)
  Dim theRect
  Dim thePos
  theRect = mDoc.Rect
  mDoc.Rect.Inset Padding, Padding
  AddText = AddToRow(mDoc.AddText(inText))
  If mTruncated = False Then
    Dim theDrawn
    theDrawn = mDoc.GetInfo(AddText, "Characters")
    If theDrawn = "" Then theDrawn = 0
    If CInt(theDrawn) < Len(inText) Then
      'mTruncated = True
    End If
  End If
  thePos = mDoc.Pos.Y - mDoc.FontSize
  If thePos < mRowBottom Then mRowBottom = thePos
  mDoc.Rect = theRect
End Function

' Select the entire table area
Public Sub SelectTable()
  mDoc.Rect = mRect
End Sub

' Select a cell in the current row using a zero based index
Public Sub SelectCell(inIndex)
  GetRowHeight mYPos ' fix the current row height
  SelectCells inIndex, mYPos, inIndex, mYPos
End Sub

' Select a row on the current page using a zero based index
Public Sub SelectRow(inIndex)
  GetRowHeight mYPos ' fix the current row height
  SelectCells 0, inIndex, mColumns - 1, inIndex
End Sub

' Select a column on the current page using a zero based index
Public Sub SelectColumn(inIndex)
  GetRowHeight mYPos ' fix the current row height
  SelectCells inIndex, 0, inIndex, UBound(mHeights) - 1
End Sub

' Select a rectangular area of cells on the current page
Public Sub SelectCells(inX1, inY1, inX2, inY2)
  Dim theTop, theLeft, theTemp, i
  ' check inputs
  If inX1 > inX2 Then
    theTemp = inX1
    inX1 = inX2
    inX2 = theTemp
  End If
  If inY1 > inY2 Then
    theTemp = inX1
    inY1 = inY2
    inY2 = theTemp
  End If
  GetRowHeight mYPos ' fix the current row height
  If inY1 >= UBound(mHeights) Then inY1 = UBound(mHeights) - 1
  If inY2 >= UBound(mHeights) Then inY2 = UBound(mHeights) - 1
  If inY1 < 0 Then Exit Sub
  ' select the cells
  mDoc.Rect = mRect
  theTop = mDoc.Rect.Top
  SelectCurrentCell inX1
  theLeft = mDoc.Rect.Left
  SelectCurrentCell inX2
  mDoc.Rect.Top = theTop
  mDoc.Rect.Bottom = theTop
  mDoc.Rect.Left = theLeft
  For i = 0 To inY2
    mDoc.Rect.Bottom = mDoc.Rect.Bottom - mHeights(i)
    If inY1 > i Then mDoc.Rect.Top = mDoc.Rect.Top - mHeights(i)
  Next
End Sub

' Draw borders round the current selection
Public Sub Frame(inTop, inBott, inLeft, inRight)
  If inTop Then AddToRow mDoc.AddLine(mDoc.Rect.Left, mDoc.Rect.Top, mDoc.Rect.Right, mDoc.Rect.Top)
  If inBott Then AddToRow mDoc.AddLine(mDoc.Rect.Left, mDoc.Rect.Bottom, mDoc.Rect.Right, mDoc.Rect.Bottom)
  If inLeft Then AddToRow mDoc.AddLine(mDoc.Rect.Left, mDoc.Rect.Top, mDoc.Rect.Left, mDoc.Rect.Bottom)
  If inRight Then AddToRow mDoc.AddLine(mDoc.Rect.Right, mDoc.Rect.Top, mDoc.Rect.Right, mDoc.Rect.Bottom)
End Sub

' Color the background of the current selection
Public Sub Fill(inColor)
  Dim theLayer, theColor
  theLayer = mDoc.Layer
  theColor = mDoc.Color
  mDoc.Layer = mDoc.LayerCount + 1
  mDoc.Color = inColor
  AddToRow mDoc.FillRect
  mDoc.Color = theColor
  mDoc.Layer = theLayer
End Sub

' Get the current row height based on the cell contents drawn so far
Private Function GetRowHeight(inRow)
  If UBound(mHeights) <= inRow Then
    ' establish and store current row height
    Dim theHeight
    ReDim Preserve mHeights(inRow + 1)
    mRowBottom = mRowBottom - Padding
    If mRowBottom < mRectTop - mRectHeight Then mRowBottom = mRectTop - mRectHeight
    theHeight = mRowTop - mRowBottom
    If inRow > 0 And theHeight < RowHeightMin Then theHeight = RowHeightMin
    If RowHeightMax > 0 And theHeight > RowHeightMax Then theHeight = RowHeightMax
    mHeights(inRow) = theHeight
  End If
  GetRowHeight = 0
  If inRow >= 0 Then GetRowHeight = mHeights(inRow)
End Function

' Select the current cell
Private Sub SelectCurrentCell(inIndex)
  Dim thePos, theWidth, theTotal, i
  If inIndex >= 0 And inIndex < mColumns Then
    ' get the x offset and width of the cell
    For i = 0 To UBound(mWidths)
      theTotal = theTotal + mWidths(i)
      If i < inIndex Then thePos = thePos + mWidths(i)
    Next
    thePos = thePos * (mRectWidth / theTotal)
    theWidth = mWidths(inIndex) * (mRectWidth / theTotal)
    ' position the cell
    mDoc.Rect.Top = mRowTop
    mDoc.Rect.Left = mRectLeft + thePos
    mDoc.Rect.Width = theWidth
  End If
End Sub

' Add to our list of objects drawn as part of the current row
Private Function AddToRow(inID)
  mObjects = mObjects &"," &inID
  AddToRow = inID
End Function

' Delete all the objects drawn as part of the current row
Public Sub DeleteLastRow()
  Dim theArray
  Dim i
  theArray = Split(mObjects, ",")
  For i = 0 To UBound(theArray)
    If theArray(i) <> "" Then mDoc.Delete CInt(theArray(i))
  Next
  mObjects = ""
End Sub

End Class ' only needed in ASP

%>