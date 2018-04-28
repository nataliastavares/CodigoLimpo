'Classe utilizada para converter 
'numeros informados em extenso

Public Class Conversor

    'Função que converte números entre 0 e 19
    Public Function Numeros0e19(ByVal dNumero As Decimal) As String
        Dim sNumeroExtenso As String = ""

        Select Case dNumero

            Case 0
                sNumeroExtenso = "zero"
            Case 1
                sNumeroExtenso = "um"
            Case 2
                sNumeroExtenso = "dois"
            Case 3
                sNumeroExtenso = "três"
            Case 4
                sNumeroExtenso = "quatro"
            Case 5
                sNumeroExtenso = "cinco"
            Case 6
                sNumeroExtenso = "seis"
            Case 7
                sNumeroExtenso = "sete"
            Case 8
                sNumeroExtenso = "oito"
            Case 9
                sNumeroExtenso = "nove"
            Case 10
                sNumeroExtenso = "dez"
            Case 11
                sNumeroExtenso = "onze"
            Case 12
                sNumeroExtenso = "doze"
            Case 13
                sNumeroExtenso = "treze"
            Case 14
                sNumeroExtenso = "quatorze"
            Case 15
                sNumeroExtenso = "quinze"
            Case 16
                sNumeroExtenso = "dezesseis"
            Case 17
                sNumeroExtenso = "dezessete"
            Case 18
                sNumeroExtenso = "dezoito"
            Case 19
                sNumeroExtenso = "dezenove"
        End Select
        Return sNumeroExtenso
    End Function


End Class
