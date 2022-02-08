Function myreverse(ByVal S As String) As String
'renvoie: renverse une chaine de caractere
Dim n As Long 'le nombre de caracteres de la chaine
n = Len(S)
If n > 1 Then
    myreverse = Right(S, 1) & myreverse(Left(S, n - 1))
Else
    myreverse = S
End If
End Function
