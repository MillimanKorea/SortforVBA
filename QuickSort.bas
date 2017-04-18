Public Sub QuickSort(ByRef SrcArray() As String, ByVal min As Long, ByVal max As Long)
    
    Dim med_value As String
    Dim hi As Long
    Dim lo As Long
    Dim i As Long
    Dim j As Long, k As Long
        
    If max <= min Then Exit Sub
    
    'i = Int((max - min + 1) * 0.5 + min)
	i = Int((max + min + 1) / 2)
    med_value = SrcArray(i)
    SrcArray(i) = SrcArray(min)
    
    lo = min
    hi = max
    
    For j = min To max
        
        For k = hi To lo Step -1
            If SrcArray(k) < med_value Or k <= lo Then
                hi = k
                Exit For
            End If
        Next k
     
        If hi <= lo Then
            SrcArray(lo) = med_value
            Exit For
        End If
    
        SrcArray(lo) = SrcArray(hi)
        lo = lo + 1
     
        For k = lo To hi
            If SrcArray(lo) >= med_value Or lo >= hi Then
                lo = k
                Exit For
            End If
        Next k
     
        If lo >= hi Then
            lo = hi
            SrcArray(hi) = med_value
            Exit For
        End If
     
        ' Swap the lo and hi values.
        SrcArray(hi) = SrcArray(lo)
    Next j
    
	QSCount = QSCount + 1
	'Debug.Print QSCount, min, max, hi, lo, SrcArray(hi), SrcArray(lo)
	
    Call QuickSort(SrcArray, min, lo - 1)
    Call QuickSort(SrcArray, lo + 1, max)

End Sub
