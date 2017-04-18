Sub InsertionSort(ByRef varData() As String, ByVal the_end As Long)

    Dim lngCounter1 As Long
    Dim lngCounter2 As Long
    Dim varTemp As String

    For lngCounter1 = 1 To the_end
        varTemp = varData(lngCounter1)
        For lngCounter2 = lngCounter1 To 1 Step -1
            If varData(lngCounter2 - 1) > varTemp Then
                varData(lngCounter2) = varData(lngCounter2 - 1)
            Else
                Exit For
            End If
        Next lngCounter2
        varData(lngCounter2) = varTemp
    Next lngCounter1

End Sub
