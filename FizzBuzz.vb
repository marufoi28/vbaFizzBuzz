Public wsExercise As Worksheet
Public wsProblem As Worksheet
Public wsTest As Worksheet
Public FizzCount As Integer
Public BuzzCount As Integer
Public FizzBuzzCount As Integer

Private Sub InitializeWorksheet()
    Set wsExercise = ThisWorkbook.Sheets("演習")
    Set wsProblem = ThisWorkbook.Sheets("問題")
    Set wsTest = ThisWorkbook.Sheets("test")
    FizzCount = 0
    BuzzCount = 0
    FizzBuzzCount = 0
End Sub

Public Sub outputFizzBuzz()
  Call InitializeWorksheet
  Dim row As Integer: row = 2
  Dim NumberListLastRow As Long: NumberListLastRow = wsExercise.Cells(wsExercise.Rows.Count, 1).End(xlUp).row

  Call clearFizzBuzz

  Do While row <= NumberListLastRow

      Dim inputNum As String: inputNum = wsExercise.Cells(row, 1)

      If Not RunFizzBuzz(inputNum) Then
        Exit Do
      End If

      Dim num As Integer: num = CInt(inputNum)

      Call writeFizzBuzz(num, row)

      row = row + 1
  Loop

  wsExercise.Cells(2, 4) = FizzCount
  wsExercise.Cells(2, 5) = BuzzCount
  wsExercise.Cells(2, 6) = FizzBuzzCount

End Sub

Private Function RunFizzBuzz(ByVal inputNum As String) As Boolean
    RunFizzBuzz = True
    If inputNum = "" Or inputNum = "end" Then
        RunFizzBuzz = False
    End If


End Function

Public Sub writeFizzBuzz(ByVal num As Long, ByVal row As Integer)
      If num Mod 15 = 0 Then
          wsExercise.Cells(row, 2) = "FizzBuzz"
          FizzBuzzCount = FizzBuzzCount + 1
      ElseIf num Mod 3 = 0 Then
        wsExercise.Cells(row, 2) = "Fizz"
        FizzCount = FizzCount + 1
     ElseIf num Mod 5 = 0 Then
       wsExercise.Cells(row, 2) = "Buzz"
       BuzzCount = BuzzCount + 1
     Else
         wsExercise.Cells(row, 2) = num
      End If
End Sub

Public Sub test02()
  Call InitializeWorksheet
  Dim row As Long: row = 2

  'TODO: NULL判定
  row = wsProblem.Cells(9, 2).Value

  Dim NumberListLastRow As Long: NumberListLastRow = wsExercise.Cells(wsExercise.Rows.Count, 1).End(xlUp).row

  Call clearFizzBuzz

    Do While row <= NumberListLastRow

      Dim inputNum As String: inputNum = wsExercise.Cells(row, 1)

      If Not RunFizzBuzz(inputNum) Then
        Exit Do
      Else
        Dim num As Integer: num = CInt(inputNum)
      End If

      writeFizzBuzz (num)
      row = row + 1
  Loop

  wsExercise.Cells(2, 4) = FizzCount
  wsExercise.Cells(2, 5) = BuzzCount
  wsExercise.Cells(2, 6) = FizzBuzzCount

End Sub

Public Sub isPrimeCheck()
    Call InitializeWorksheet
    Dim row As Long: row = 2
    Dim NumberListLastRow As Long: NumberListLastRow = wsExercise.Cells(wsExercise.Rows.Count, 1).End(xlUp).row
    Do While row <= NumberListLastRow
        Dim inputNum As String: inputNum = wsExercise.Cells(row, 1)
        If inputNum = "" Or inputNum = "end" Then
          Exit Do
        Else
          num = wsExercise.Cells(row, 1)
        End If

        If isPrime(wsExercise.Cells(row, 1)) Then
            wsExercise.Cells(row, 3) = "●"
        End If
        row = row + 1
    Loop
End Sub

'素数判定
'素数の定義は「自然数の中で1とその数自身以外に約数が無いもの」
Function isPrime(n As Long) As Boolean
    Dim i As Long

    If n <= 1 Then
      isPrime = False
      Exit Function
    End If

    If n Mod 2 = 0 Or n Mod 3 = 0 Then
        isPrime = False
        Exit Function
    End If

    For i = 5 To Sqr(n) Step 2
        If n Mod i = 0 Then
            isPrime = False
            Exit Function
        End If
    Next i

    isPrime = True
End Function

'クリアボタン
Public Sub clear()
    Dim row As Integer: row = 2
    Call InitializeWorksheet
    Call clearFizzBuzz
    Call clearIsPrimeNumber
End Sub

'Fizz/Buzzのクリア
Private Sub clearFizzBuzz()
    Call InitializeWorksheet
    Dim fizzBuzzLastRow As Long: fizzBuzzLastRow = wsExercise.Cells(wsExercise.Rows.Count, 2).End(xlUp).row
    Dim row As Long: row = 2

    Do While row <= fizzBuzzLastRow
        wsExercise.Cells(row, 2).clear
    row = row + 1
    Loop
End Sub

'PrimeNumberのクリア
Private Sub clearIsPrimeNumber()
    Call InitializeWorksheet
    Dim isPrimeNumberLastRow As Long: isPrimeNumberLastRow = wsExercise.Cells(wsExercise.Rows.Count, 3).End(xlUp).row
    Dim row As Long: row = 2

    Do While row <= isPrimeNumberLastRow
        wsExercise.Cells(row, 3).clear
    row = row + 1
    Loop
End Sub
