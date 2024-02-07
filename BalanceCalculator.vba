' BalanceCalculator osztály
Public monthName As String
Public rowNumber As Integer
Private Const Months As String = "Január,Február,Március,Április,Május,Június,Július,Augusztus,Szeptember,Október,November,December"

Public Function CurrentBalance() As Variant
  If Not IsValidMonth() Then
    Err.Raise 5, , "Érvénytelen hónap név"
  End If
  
  If Not IsValidRow() Then
    Err.Raise 5, , "Érvénytelen sor szám"
  End If
  
  Dim amounts() As Variant
  amounts = GetValuesFromTable(GetColumnNumber())
  
  CurrentBalance = SumFromFirstNegative(amounts)
End Function

Private Function IsValidMonth() As Boolean
  ' A hónap neve érvényes-e
  IsValidMonth = InStr(1, Months, monthName) > 0
End Function

Private Function IsValidRow() As Boolean
  ' A sor száma érvényes-e
  IsValidRow = rowNumber > 0 And rowNumber <= ThisWorkbook.Sheets(1).Rows.Count
End Function

Private Function GetColumnNumber() As Integer
  Dim rng As Range
  Set rng = ThisWorkbook.Sheets(1).Rows(3).Find(monthName, LookIn:=xlValues)
  
  If rng Is Nothing Then
    Err.Raise 5, , "A hónap nem található"
  End If
  
  GetColumnNumber = rng.Column
End Function

Private Function GetValuesFromTable(columnNumber As Integer) As Variant
  ' Egy tömb az értékek tárolására
  Dim values() As Variant
  ReDim values(1 To columnNumber - 2)
  
  ' Beolvassuk az értékeket a tömbbe
  For i = 3 To columnNumber
    values(i - 2) = ThisWorkbook.Sheets(1).Cells(rowNumber, i).Value
  Next i
  
  GetValuesFromTable = values
End Function

Private Function SumFromFirstNegative(values() As Variant) As Variant
  Dim sum As Variant
  Dim foundNegative As Boolean
  sum = 0
  
  foundNegative = False

  For i = LBound(values) To UBound(values)
    If values(i) < 0 Then
      foundNegative = True
    End If
   
    If foundNegative Then
      sum = sum + values(i)
      If sum > 0 And i <> UBound(values) Then
        sum = 0
      End If
    End If
  Next i
  
  If sum = 0 And Not foundNegative Then
    SumFromFirstNegative = values(UBound(values))
  'ElseIf sum = 0 And foundNegative Then
  '  SumFromFirstNegative = 0
  'ElseIf sum < 0 Then
  '  SumFromFirstNegative = 0
  Else
    SumFromFirstNegative = sum
  End If
End Function
