Private Sub GetRunTimes1()
    Dim i&, ii&
    Dim RunRow&
    For i = 2 To 999999
        If Cells(i, 1) = 0 Then Exit Sub
        If Cells(i, 5) = "Run" Then
            RunRow = i
            For ii = RunRow + 1 To 999999
                If Cells(ii, 1) = 0 Then Exit Sub
                If Cells(ii, 5) = "Off" Then
                    Cells(ii, 34) = (Cells(ii, 2) - Cells(RunRow, 2)) * 1440
                    Exit For
                End If
            Next
        End If
    Next
End Sub


Private Sub GetRunTimes2()
    Dim i&, ii&
    Dim RunRow&
    For i = 2 To 999999
        If Cells(i, 1) = 0 Then Exit Sub
        If Cells(i, 9) = "Run" Then
            RunRow = i
            For ii = RunRow + 1 To 999999
                If Cells(ii, 1) = 0 Then Exit Sub
                If Cells(ii, 9) = "Off" Then
                    Cells(ii, 35) = (Cells(ii, 2) - Cells(RunRow, 2)) * 1440
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Private Sub GetRunTimes3()
    Dim i&, ii&
    Dim RunRow&
    For i = 2 To 999999
        If Cells(i, 1) = 0 Then Exit Sub
        If Cells(i, 13) = "Run" Then
            RunRow = i
            For ii = RunRow + 1 To 999999
                If Cells(ii, 1) = 0 Then Exit Sub
                If Cells(ii, 13) = "Off" Then
                    Cells(ii, 36) = (Cells(ii, 2) - Cells(RunRow, 2)) * 1440
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Private Sub GetRunTimes4()
    Dim i&, ii&
    Dim RunRow&
    For i = 2 To 999999
        If Cells(i, 1) = 0 Then Exit Sub
        If Cells(i, 17) = "Run" Then
            RunRow = i
            For ii = RunRow + 1 To 999999
                If Cells(ii, 1) = 0 Then Exit Sub
                If Cells(ii, 17) = "Off" Then
                    Cells(ii, 37) = (Cells(ii, 2) - Cells(RunRow, 2)) * 1440
                    Exit For
                End If
            Next
        End If
    Next
End Sub

Sub TransferRun()
Worksheets("DATA").Range("AH2:AH99999").Copy Worksheets("DLRT").Range("C2:C99999")
Worksheets("DATA").Range("AI2:Ai99999").Copy Worksheets("DLRT").Range("E2:E99999")
Worksheets("DATA").Range("AJ2:AJ99999").Copy Worksheets("DLRT").Range("G2:G99999")
Worksheets("DATA").Range("AK2:AK99999").Copy Worksheets("DLRT").Range("I2:I99999")

End Sub

Sub RemoveBlankCells()

Dim rng As Range

  On Error GoTo NoBlanksFound
    Set rng = Worksheets("DLRT").Range("B2:I2000").SpecialCells(xlCellTypeBlanks)
  On Error GoTo 0

  rng.Rows.Delete Shift:=xlShiftUp

Exit Sub

NoBlanksFound:
  MsgBox "No Blank cells were found"

End Sub

Sub ClearTable()

Worksheets("DATA").Range("AH2:AK99999").ClearContents

End Sub

