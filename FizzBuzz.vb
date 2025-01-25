Public wsExercise As Worksheet
Public wsProblem As Worksheet
Public wsTest As Worksheet

Public Sub button1_Click()
    outputFizzBuzz 1
End Sub

Public Sub button2_Click()
    outputFizzBuzz 2
End Sub

Public Sub outputFizzBuzz(buttonType As Integer)
    Dim fizzCount As Integer, buzzCount As Integer, fizzBuzzCount As Integer
    fizzCount = 0
    buzzCount = 0
    fizzBuzzCount = 0
    Set wsExercise = ThisWorkbook.Sheets("演習")
    Set wsProblem = ThisWorkbook.Sheets("問題")

    Dim row As Long
    If buttonType = 1 Then
        row = 2
    ElseIf buttonType = 2 Then
        row = wsProblem.Cells(9, 2)
    End If
    Dim NumberListLastRow As Long : NumberListLastRow = wsExercise.Cells(wsExercise.Rows.Count, 1).End(xlUp).row

    clearColumn 2

    Do While row <= NumberListLastRow

        Dim inputNum As String : inputNum = wsExercise.Cells(row, 1)

        If inputNum = "" Or inputNum = "end" Then Exit Do

        If IsNumeric(inputNum) Then
            Dim num As Integer : num = CLng(inputNum)
            writeFizzBuzz num, row, fizzCount, buzzCount, fizzBuzzCount
        End If
        row = row + 1
    Loop

    wsExercise.Cells(2, 4) = fizzCount
    wsExercise.Cells(2, 5) = buzzCount
    wsExercise.Cells(2, 6) = fizzBuzzCount

End Sub

Public Sub writeFizzBuzz(ByVal num As Long, ByVal row As Long,
                         ByRef fizzCount As Integer, ByRef buzzCount As Integer, ByRef fizzBuzzCount As Integer)
    If num Mod 15 = 0 Then
        wsExercise.Cells(row, 2) = "FizzBuzz"
        fizzBuzzCount = fizzBuzzCount + 1
    ElseIf num Mod 3 = 0 Then
        wsExercise.Cells(row, 2) = "Fizz"
        fizzCount = fizzCount + 1
    ElseIf num Mod 5 = 0 Then
        wsExercise.Cells(row, 2) = "Buzz"
        buzzCount = buzzCount + 1
    Else
        wsExercise.Cells(row, 2) = num
    End If
End Sub

Private Sub isPrimeCheck()
    Dim row As Long : row = 2
    Dim NumberListLastRow As Long : NumberListLastRow = wsExercise.Cells(wsExercise.Rows.Count, 1).End(xlUp).row
    Do While row <= NumberListLastRow
        Dim inputNum As String : inputNum = wsExercise.Cells(row, 1)
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

    For i = 2 To Sqr(n)
        If n Mod i = 0 Then
            isPrime = False
            Exit Function
        End If
    Next i

    isPrime = True
End Function

'クリアボタン
Public Sub clear()
    clearColumn 2
    clearColumn 3
End Sub

'行のクリア
Private Sub clearColumn(columNumber As Long)
    Set wsExercise = ThisWorkbook.Sheets("演習")
    Dim lastRow As Long : lastRow = wsExercise.Cells(wsExercise.Rows.Count, columNumber).End(xlUp).row
    Dim row As Long : row = 2

    Do While row <= lastRow
        wsExercise.Cells(row, columNumber).clear
        row = row + 1
    Loop
End Sub



