Sub MatriceToList()
'
' MatriceToList Macro
'

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

'Get A Cell Address From The User to Get Number Format From
  On Error Resume Next
    Set rng = Application.InputBox( _
      Title:="Transformer Matrice en Liste", _
      prompt:="Selectionnez les données de la Matrice", _
      Default:=DefaultRange.Address, _
      Type:=8)
  On Error GoTo 0

'Test to ensure User Did not cancel
  If rng Is Nothing Then Exit Sub
  
  selectedRange = rng.Address
  'MsgBox (selectedRange + "Type : " + TypeName(selectedRange))
  stringSplitted = Split(selectedRange, ":")
  
  'For Each stringItem In stringSplitted
    'MsgBox (stringItem)
  'Next
  
  firstCell = Replace(stringSplitted(0), "$", "")
  firstCellXY(0) = Range(firstCell).Row
  firstCellXY(1) = Range(firstCell).Column
  'MsgBox ("First Cell : " + firstCell + "-  X : " + CStr(firstCellXY(0)) + " Y : " + CStr(firstCellXY(1)))
  
  
  
  lastCell = Replace(stringSplitted(1), "$", "")
  lastCellXY(0) = Range(lastCell).Row
  lastCellXY(1) = Range(lastCell).Column
  'MsgBox ("Last Cell : " + lastCell + "-  X : " + CStr(lastCellXY(0)) + " Y : " + CStr(lastCellXY(1)))
  
  numberOfRows = rng.Rows.Count
  numberOfColumns = rng.Columns.Count
  
  'MsgBox ("Nombre de Colonnes sélectionnées : " + CStr(numberOfColumns))
  'MsgBox ("Nombre de Lignes sélectionnées : " + CStr(numberOfRows))

  If (numberOfRows = numberOfColumns) Then
    MsgBox ("La Matrice est bien carré !")
  Else
    MsgBox ("La Matrice non carré ! Arrêt en cours ...")
    Exit Sub
  End If
  
  Set startOfTableCell = Application.InputBox(prompt:="Please select any cell", Type:=8)
  startOfTableCell = Replace(startOfTableCell.Address, "$", "")
  startOfTableCellXY(0) = Range(startOfTableCell).Row
  startOfTableCellXY(1) = Range(startOfTableCell).Column
  'MsgBox (startOfTableCell + " - X : " + CStr(startOfTableCellXY(0)) + " Y : " + CStr(startOfTableCellXY(1)))
  

  j = 1
  k = 1
  cptJ = 0
  
  currentCellX = firstCellXY(0)
  currentCellY = firstCellXY(1)
  
  For i = 1 To (numberOfColumns * numberOfRows)
    Cells(startOfTableCellXY(0) + (i - 1), startOfTableCellXY(1)) = j
    Cells(startOfTableCellXY(0) + (i - 1), startOfTableCellXY(1) + 1) = k
    
    Cells(startOfTableCellXY(0) + (i - 1), startOfTableCellXY(1) + 2) = Cells(currentCellX, currentCellY).Value
    currentCellY = currentCellY + 1
    cptJ = cptJ + 1
    k = k + 1

    If (cptJ = numberOfColumns) Then
        j = j + 1
        cptJ = 0
        k = 1
        currentCellX = currentCellX + 1
        currentCellY = firstCellXY(1)
    End If
    
  Next

'Highlight Cell Range
  rng.Interior.Color = vbYellow

End Sub
