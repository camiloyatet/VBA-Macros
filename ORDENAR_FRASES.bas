'This macro splits a phrase and sort each word

Public Function ORDENAR_FRASES(Txt_Entrada As String, Descendente As Boolean, Delimitador As String) As String

Dim VerPalabra() As String
Dim text_string As String
VerPalabra() = Split(Txt_Entrada, Delimitador)

If Descendente Then

For x = LBound(VerPalabra) To UBound(VerPalabra)
    For y = x To UBound(VerPalabra)
        If UCase(VerPalabra(y)) < UCase(VerPalabra(x)) Then
            TempTxt1 = VerPalabra(x)
            TempTxt2 = VerPalabra(y)
            VerPalabra(x) = TempTxt2
            VerPalabra(y) = TempTxt1
        End If
    Next y
Next x

Else

For x = LBound(VerPalabra) To UBound(VerPalabra)
    For y = x To UBound(VerPalabra)
        If UCase(VerPalabra(y)) > UCase(VerPalabra(x)) Then
            TempTxt1 = VerPalabra(x)
            TempTxt2 = VerPalabra(y)
            VerPalabra(x) = TempTxt2
            VerPalabra(y) = TempTxt1
        End If
    Next y
Next x

End If

For i = LBound(VerPalabra) To UBound(VerPalabra)
  text_string = text_string & " " & VerPalabra(i)
Next i

ORDENAR_FRASES = UCase(Trim(text_string))

End Function
