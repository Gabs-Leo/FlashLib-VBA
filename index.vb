'Comparacoes
'Confere se um número está entre dois números
Function isBetween(low, number, high)
    isBetween = low < number And number < high
End Function

'Confere se um número está entre dois números, considerando os dois números
Function isFromTo(low, number, high)
    isFromTo = low <= number And number <= high
End Function
'-----------------------------------------------------------
'Utilitarios para Arrays
'Retorna o tamanho de um array
Function getSize(someArray) As Integer
    getSize = UBound(someArray) - LBound(someArray) + 1
End Function

'Adiciona um novo índice em um array
Function push(someArray, someValue)
    arraySize = getSize(someArray)
    ReDim Preserve someArray(arraySize)
    someArray(arraySize) = someValue
    push = someArray
End Function

'-----------------------------------------------------------
'Utilitarios para Console
'Recebe um array e retorna uma string com quebras de linhas pra cada elemento no array
Function writeBreak(stringArray)
    myS = ""
    For i = 0 To getSize(stringArray) - 1
        myS = myS & stringArray(i) & vbCrLf
    Next i
    writeBreak = myS
End Function
