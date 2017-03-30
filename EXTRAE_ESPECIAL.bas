' This function allows to clean a string and returns a letters only, number only or alphanumeric only string 
' according to Tipo parameter. It is possible to include or exclude spaces with the boolean parameter: Espacios

Public Enum Extraccion
    Alfabetico = 1
    Numerico = 2
    Alfanumerico = 3
End Enum

Public Function EXTRAE_ESPECIAL(Txt_Entrada As String, Tipo As Extraccion, Espacios As Boolean) As String

Dim objRegEx As Object
Dim strPattern As String
Set objRegEx = CreateObject("vbscript.regexp")

If Espacios Then

    Select Case Tipo
        Case Extraccion.Alfabetico
            strPattern = "[^a-zA-Z]+"
        Case Extraccion.Numerico
            strPattern = "[^0-9]+"
        Case Extraccion.Alfanumerico
            strPattern = "[^a-zA-Z0-9]+"
    End Select
    
Else

    Select Case Tipo
        Case Extraccion.Alfabetico
            strPattern = "[^a-zA-Z\s]+"
        Case Extraccion.Numerico
            strPattern = "[^0-9\s\s]+"
        Case Extraccion.Alfanumerico
            strPattern = "[^a-zA-Z0-9\s]+"
    End Select

End If

    With objRegEx
    .Global = True
    .Pattern = strPattern
    GetDrinkSpecial = .Replace(Replace(Txt_Entrada, "-", Chr(32)), vbNullString)
    End With

End Function
