Public Sub HeapSort(ByRef list() As String, ByVal Ascending As Boolean, ByVal k As Long)
 
    Dim j As String, the_end As Long
 
    Call heapify(list, k, Ascending)
 
    the_end = k
    Do While the_end >= 1
        j = list(the_end)
        list(the_end) = list(1)
        list(1) = j

        the_end = the_end - 1

        Call siftDown(list, 1, the_end, Ascending)
    Loop
 
End Sub


Public Sub heapify(ByRef list() As String, ByVal count As Long, ByVal Ascending As Boolean)
 
    Dim start As Long

    'start = (count - 2) / 2
	start = count / 2
      
    Do While start > 0
        'Call siftDown(list, start, count - 1, Ascending)
		Call siftDown(list, start, count, Ascending)
        start = start - 1
    Loop

End Sub


Public Sub siftDown(ByRef list() As String, ByVal the_start As Long, ByVal the_end As Long, ByVal Ascending As Boolean)
 
    Dim k As String, root As Long, child As Long, swap As Long
    root = the_start
 
    Do While root * 2 <= the_end
        child = root * 2
        swap = root
  
        If Ascending Then
            If list(swap) < list(child) Then
                swap = child
            End If
  
            If child + 1 <= the_end And list(swap) < list(child + 1) Then
                swap = child + 1
            End If
        Else
  
            If list(child) < list(swap) Then
                swap = child
            End If
   
            If child + 1 <= the_end And list(child + 1) < list(swap) Then
                swap = child + 1
            End If
        End If
          
        If swap <> root Then
            k = list(root)
            list(root) = list(swap)
            list(swap) = k
            root = swap
        Else
            Exit Sub
        End If
 
    Loop
 
End Sub
