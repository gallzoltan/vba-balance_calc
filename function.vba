Function CurrentBalance(monthName As String, rowNumber As Integer) As Variant
  If Not IsValidMonth(monthName) Then
    CurrentBalance = "Érvénytelen hónap név"
    Exit Function
  End If
  
  If Not IsValidRow(rowNumber) Then
    CurrentBalance = "Érvénytelen sor szám"
    Exit Function
  End If
  
  Dim columnNumber As Integer
  columnNumber = GetColumnNumber(monthName)

  Dim values() As Variant
  values = GetValuesFromTable(rowNumber, columnNumber)
  
  Dim sum As Variant
  sum = SumFromFirstNegative(values)
  
  CurrentBalance = sum
End Function

Function IsValidMonth(monthName As String) As Boolean
  ' A hónap neve érvényes-e
  Select Case monthName
    Case "Január", "Február", "Március", "Április", "Május", "Június", _
         "Július", "Augusztus", "Szeptember", "Október", "November", "December"
      IsValidMonth = True
    Case Else
      IsValidMonth = False
  End Select
End Function

Function IsValidRow(rowNumber As Integer) As Boolean
  ' A sor száma érvényes-e
  If rowNumber > 0 And rowNumber <= ThisWorkbook.Sheets(1).Rows.Count Then
    IsValidRow = True
  Else
    IsValidRow = False
  End If
End Function

Function GetColumnNumber(monthName As String) As Integer
  Dim rng As Range
  Set rng = ThisWorkbook.Sheets(1).Rows(3).Find(monthName, LookIn:=xlValues)
  
  If Not rng Is Nothing Then
    GetColumnNumber = rng.Column
  Else
    GetColumnNumber = 0
  End If
End Function

Function GetValuesFromTable(rowNumber As Integer, columnNumber As Integer) As Variant
  ' Egy tömb az értékek tárolására
  Dim values() As Variant
  Dim numValues As Integer
  numValues = columnNumber - 2
  ReDim values(1 To numValues)
  
  ' Beolvassuk az értékeket a tömbbe
  For i = 3 To columnNumber
    values(i - 2) = ThisWorkbook.Sheets(1).Cells(rowNumber, i).Value
  Next i
  
  GetValuesFromTable = values
End Function

Function SumFromFirstNegative(values() As Variant) As Variant
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
  ElseIf sum = 0 And foundNegative Then
    SumFromFirstNegative = 0
  ElseIf sum < 0 Then
    SumFromFirstNegative = 0
  Else
    SumFromFirstNegative = sum
  End If
  
End Function
