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
            Case 20
                sNumeroExtenso = "vinte"
            Case 30
                sNumeroExtenso = "trinta"
            Case 40
                sNumeroExtenso = "quarenta"
            Case 50
                sNumeroExtenso = "cinquenta"
            Case 60
                sNumeroExtenso = "sessenta"
            Case 70
                sNumeroExtenso = "setenta"
            Case 80
                sNumeroExtenso = "oitenta"
            Case 90
                sNumeroExtenso = "noventa"
            Case 100
                sNumeroExtenso = "cento"
            Case 200
                sNumeroExtenso = "duzentos"
            Case 300
                sNumeroExtenso = "trezentos"
            Case 400
                sNumeroExtenso = "quatrocentos"
            Case 500
                sNumeroExtenso = "quinhetos"
            Case 600
                sNumeroExtenso = "seiscentos"
            Case 700
                sNumeroExtenso = "setecentos"
            Case 800
                sNumeroExtenso = "oitocentos"
            Case 900
                sNumeroExtenso = "novecentos"
        End Select
        Return sNumeroExtenso
    End Function

    Public Function validarValoresReais(ByVal dNumero As String) As String
        Dim iDigitos As Int32
        Dim sUnidade As String
        Dim sDezenas As String
        Dim sCentenas As String
        Dim sMilhares As String
        Dim sNumeroExtenso As String = ""

        iDigitos = dNumero.ToString.Length()

        If dNumero > 19 Then
            If iDigitos = 1 Then
                sUnidade = dNumero.ToString.Substring(iDigitos - 1, 1)
                sNumeroExtenso = Me.Numeros0e19(Convert.ToInt32(sUnidade))
            ElseIf iDigitos = 2 Then
                sDezenas = dNumero.ToString.Substring(iDigitos - 2, 1)
                sNumeroExtenso = Me.Numeros0e19(Convert.ToInt32(sDezenas & "0"))
                If sUnidade <> 0 Then
                    sUnidade = dNumero.ToString.Substring(iDigitos - 1, 1)
                End If
                sNumeroExtenso &= " e " & Me.Numeros0e19(Convert.ToInt32(sUnidade))
            ElseIf iDigitos = 3 Then
                sCentenas = dNumero.ToString.Substring(iDigitos - 3, 1)
                sNumeroExtenso = Me.Numeros0e19(Convert.ToInt32(sCentenas & "00"))
                sDezenas = dNumero.ToString.Substring(iDigitos - 2, 1)
                If sDezenas <> 0 Then
                    sNumeroExtenso &= " e " & Me.Numeros0e19(Convert.ToInt32(sDezenas & "0"))
                End If
                sUnidade = dNumero.ToString.Substring(iDigitos - 1, 1)
                If sUnidade <> 0 Then
                    sNumeroExtenso &= " e " & Me.Numeros0e19(Convert.ToInt32(sUnidade))
                End If
            ElseIf iDigitos = 4 Then
                'sUnidade = dNumero.ToString.Substring(iDigitos - 1, 1)
                'sDezenas = dNumero.ToString.Substring(iDigitos - 2, 1)
                'sCentenas = dNumero.ToString.Substring(iDigitos - 3, 1)
                'sMilhares = dNumero.ToString.Substring(iDigitos - 4, 1)
                sNumeroExtenso = "Não foi possivel converter em extenso"
                Return sNumeroExtenso
            End If
        Else
            sNumeroExtenso = Me.Numeros0e19(Convert.ToInt32(dNumero))
        End If
        

        Return sNumeroExtenso & " Reais"
    End Function

End Class
