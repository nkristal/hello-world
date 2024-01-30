Attribute VB_Name = "Module1"
Sub runTen()
    For k = 1 To 3
        LifeA
    Next k
    
End Sub

Sub LifeA()
Attribute LifeA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LifeA מאקרו
'
Dim i As Integer
Dim j As Integer

'
    For i = 2 To 25
        For j = 2 To 25
            Cells(i, j + 25) = updateCell(i, j)

        Next j
    Next i
    
    copyBuffer2
    Range("a1").Select
    
    
End Sub



Function updateCell(i, j As Integer)

    c = countNeighbours(i, j)
    
    
    If Cells(i, j) = 1 Then
        updateCell = 1
        If c < 2 Then updateCell = 0
        If c > 3 Then updateCell = 0
    Else
        updateCell = 0
        If c = 3 Then updateCell = 1
    End If
    
End Function


Function countNeighbours(i, j)
    
    countNeighbours = Cells(i - 1, j - 1) + Cells(i - 1, j) + Cells(i - 1, j + 1) _
                    + Cells(i, j - 1) + Cells(i, j + 1) _
                    + Cells(i + 1, j - 1) + Cells(i + 1, j) + Cells(i + 1, j + 1)
    'Debug.Print countNeighbours

End Function

Sub oneCell()
    Cells(13, 13) = 4
    For i = 0 To 3
        c = countNeighbours(13, 13 + i)
    Next

    
End Sub

Sub copyBuffer()
    
    Range("AA2:AX25").Select
    Range("AX2").Activate
    Selection.Copy
    Range("B2:Y25").Select
    ActiveSheet.Paste

End Sub

Sub copyBuffer2()
    Range("AA2:AX25").Copy _
      Destination:=Range("B2:Y25")
End Sub

Sub wipeBoard()

'
' LifeA מאקרו
'
Dim i As Integer
Dim j As Integer

    For i = 2 To 25
        For j = 2 To 25
            Cells(i, j) = 0

        Next j
    Next i
End Sub

