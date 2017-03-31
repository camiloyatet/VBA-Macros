' Function that computes the Levenshtein's distance and return the percentage of fuzzy matching of two strings

Public Function PCT_COINCIDENCIA(s1 As String, s2 As String)

Dim i As Integer
Dim j As Integer
Dim l1 As Integer
Dim l2 As Integer
Dim d() As Integer
Dim min1 As Integer
Dim min2 As Integer

l1 = Len(s1)
l2 = Len(s2)
ReDim d(l1, l2)
For i = 0 To l1
    d(i, 0) = i
Next
For j = 0 To l2
    d(0, j) = j
Next
For i = 1 To l1
    For j = 1 To l2
        If Mid(s1, i, 1) = Mid(s2, j, 1) Then
            d(i, j) = d(i - 1, j - 1)
        Else
            min1 = d(i - 1, j) + 1
            min2 = d(i, j - 1) + 1
            If min2 < min1 Then
                min1 = min2
            End If
            min2 = d(i - 1, j - 1) + 1
            If min2 < min1 Then
                min1 = min2
            End If
            d(i, j) = min1
        End If
    Next
Next
PCT_COINCIDENCIA = 1 - (d(l1, l2) / l1)
End Function
