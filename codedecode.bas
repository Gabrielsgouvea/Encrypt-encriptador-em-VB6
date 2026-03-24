Attribute VB_Name = "codedecode"
Public Function Encode(Data As String, Optional Depth As Integer) As String ' função de encriptação
Dim TempChar As String
Dim TempAsc As Integer
Dim NewData As String
Dim vChar As Integer
For vChar = 1 To Len(Data)
     TempChar = Mid$(Data, vChar, 1) ' lê cada caractere do texto puro
     TempAsc = Asc(TempChar) ' ... e transforma-o em valor numérico
     If Depth = 0 Then Depth = 40 ' não deixa que a senha seja 0
     If Depth > 254 Then Depth = 254
     
     If TempAsc Mod 3 = 0 Then TempAsc = TempAsc + 3
     
     TempAsc = TempAsc + Depth ' substitui pela senha
     If TempAsc > 255 Then TempAsc = TempAsc - 255 ' limite numérico máximo
     TempChar = Chr(TempAsc) ' retorna para caractere e concatena o texto cifrado
     NewData = NewData & TempChar
Next vChar

'conversão para binário
total = Len(NewData)
Dim nn As Integer
'nn = loop grande
For nn = 0 To total - 1
    recip = Mid(NewData, (total - nn), 1)
Dim resto As Single, valor As Single, bininvert As String, quantnum As Integer, numbin As String, n As Integer
Dim result As String
'resto calcula o resto da divisão
'valor é o número a ser trabalhado
'bininvert número binário invertido
'quantnum quantidade de números
'numbin número binário
'n uutilizado em loop
'result resultado final
resto = Asc(recip)
valor = Asc(recip)
bininvert = ""
Do
    resto = (resto Mod 2)
    valor = (valor \ 2)
    bininvert = bininvert & resto
    resto = valor
    valor = valor
Loop While valor > 1
bininvert = bininvert & valor
quantnum = Len(bininvert)
For n = 0 To (quantnum - 1)
    numbin = numbin & Mid(bininvert, (quantnum - n), 1)
Next n
result = result + " " + numbin
numbin = ""
Next nn
Encode = result
End Function
Public Function Decode(Data As String, Optional Depth As Integer) As String ' função de desencriptação
Dim TempChar As String
Dim TempAsc As Integer
Dim NewData As String
Dim vChar As Integer
Dim fspace As Long

'saindo de binário
fspace = 1
inicio:
For fspace = fspace To Len(Data)
     guard = Mid(Data, (fspace + 1), 1)
     If guard = " " Then
     fspace = fspace + 1
     GoTo baixo
     End If
     recip = recip + guard
     guard = ""
Next fspace
baixo:


Debug.Print guard
Dim numasc As Integer, numbin As String, quantnum As Integer, bimmult As Integer, n As Variant, result As String
'numasc valor asc
'numbin valor binário
'quantnum quantidade de números
'bimmult binario multiplicado
'n utilizado em loop
'resul resultado final
numbin = recip
recip = ""
quantnum = Len(numbin)
For n = 0 To (quantnum - 1)
    bimmult = Mid(numbin, (quantnum - n), 1)
    If bimmult = 1 Then
        numasc = numasc + (bimmult * (2 ^ n))
    End If
Next n
result = result + Chr(numasc)
numasc = 0
numbim = ""
quantnum = 0
bimmult = 1
If fspace < Len(Data) Then GoTo inicio

For vChar = 1 To Len(result)
     TempChar = Mid$(result, vChar, 1) ' a parte de leitura dos caracteres é igual à usada na encriptação
     TempAsc = Asc(TempChar)
     If Depth = 0 Then Depth = 40 ' não deixa que a senha seja 0
     If Depth > 254 Then Depth = 254
     TempAsc = TempAsc - Depth
     
     If TempAsc Mod 3 = 0 Then TempAsc = TempAsc - 3
     
     If TempAsc < 0 Then TempAsc = TempAsc + 255
     TempChar = Chr(TempAsc)
     NewData = NewData & TempChar
Next vChar
For n = 0 To (Len(NewData) - 1)
    Decode = Decode & Mid(NewData, (Len(NewData) - n), 1)
Next n
End Function
