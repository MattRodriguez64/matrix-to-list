Sub MatriceToList()
'***************************************************************************
'@Name: MatrixToList
'@Version: Final
'@Purpose: transforms the data in a matrix into a list
'@Inputs:  rng - the range selected by the user
'         startOfTableCell - cell of the future list that will be generated
'@Outputs: a list containing the X and Y coordinates and the data present at that position in the matrix
'***************************************************************************

'Variables
Dim rng As Range
Dim DefaultRange As Range
Dim FormatRuleInput As String
Dim numberOfRows As Integer
Dim numberOfColumns As Integer
Dim selectedRange As String
Dim stringSplitted() As String
Dim stringItem As Variant
Dim firstCell As String
Dim lastCell As String
Dim firstCellXY(2) As Integer
Dim lastCellXY(2) As Integer
Dim i, j, k As Integer
Dim cptJ As Integer
Dim startOfTableCellXY(2) As Integer
Dim currentCellX, currentCellY As Integer

'Determine a default range based on user's Selection
  If TypeName(Selection) = "Range" Then
    Set DefaultRange = Selection
  Else
    Set DefaultRange = ActiveCell
  End If

'Get the range selected by the user
  On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Transformer Matrice en Liste", _
      prompt:="Selectionnez les données de la Matrice", _
      Default:=DefaultRange.Address, _
      Type:=8)
  On Error GoTo 0

'Test to ensure User Did not cancel
  If rng Is Nothing Then Exit Sub
  
  'Get the range selected by the user
  selectedRange = rng.Address
  'Split the String to get rid of the colon caracter
  stringSplitted = Split(selectedRange, ":")
  
  'Replace the dollar caracter to only have the cell name
  firstCell = Replace(stringSplitted(0), "$", "")
  'Get the row number of the first cell in the selection
  firstCellXY(0) = Range(firstCell).Row
  'Get the column number of the first cell in the selection
  firstCellXY(1) = Range(firstCell).Column
  
  'Replace the dollar caracter to only have the cell name
  lastCell = Replace(stringSplitted(1), "$", "")
  'Get the row number of the last cell in the selection
  lastCellXY(0) = Range(lastCell).Row
  'Get the column number of the last cell in the selection
  lastCellXY(1) = Range(lastCell).Column
  
  'Count the number of Rows and Columns in the selection
  numberOfRows = rng.Rows.Count
  numberOfColumns = rng.Columns.Count
  
  'Check if the matrix is square
  If (numberOfRows = numberOfColumns) Then
    MsgBox ("La Matrice est bien carré !")
  Else
    MsgBox ("La Matrice non carré ! Arrêt en cours ...")
    Exit Sub
  End If
  
  'Get the cell of the future list that will be generated
  Set startOfTableCell = Application.InputBox(prompt:="Please select any cell", Type:=8)
  'Replace the dollar caracter to only have the cell name
  startOfTableCell = Replace(startOfTableCell.Address, "$", "")
  'Get the row number of the first cell of the list
  startOfTableCellXY(0) = Range(startOfTableCell).Row
  'Get the column number of the first cell of the list
  startOfTableCellXY(1) = Range(startOfTableCell).Column
  
  'j represents the rows of the matrix
  j = 1
  'k represents the columns of the matrix
  k = 1
  'cptJ is a counter
  cptJ = 0
  
  'Get the current cell in the original matrix
  currentCellX = firstCellXY(0)
  currentCellY = firstCellXY(1)
  
  'Loop to create the list row by row
  For i = 1 To (numberOfColumns * numberOfRows)
    'Writes the current row number of the matrix
    Cells(startOfTableCellXY(0) + (i - 1), startOfTableCellXY(1)) = j
    'Writes the current column number of the matrix
    Cells(startOfTableCellXY(0) + (i - 1), startOfTableCellXY(1) + 1) = k
    
    'Writes the current data present in the currentCell of the original matrix
    Cells(startOfTableCellXY(0) + (i - 1), startOfTableCellXY(1) + 2) = Cells(currentCellX, currentCellY).Value
    'Change the column
    currentCellY = currentCellY + 1
    cptJ = cptJ + 1
    'Change the column
    k = k + 1
    
    'Check if we're currently in the last cell of the matrix
    If (cptJ = numberOfColumns) Then
        'Change Row
        j = j + 1
        'Reset the counter
        cptJ = 0
        'Reset the row number
        k = 1
        'Change line
        currentCellX = currentCellX + 1
        'Go back to the first colum of the matrix
        currentCellY = firstCellXY(1)
    End If
    
  Next

  'Highlight Cell Range
  rng.Interior.Color = vbYellow

End Sub
