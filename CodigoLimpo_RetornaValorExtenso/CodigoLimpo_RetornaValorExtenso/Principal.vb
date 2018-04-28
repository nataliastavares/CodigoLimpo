Module Principal

    Sub Main()

        Dim objConversor As New Conversor
        Dim bNumero As Decimal
        Dim sNumeroExtenso As String

        'Recebe um número pelo usuário 
        Console.Write("Digite um número em reais: ")
        'Faz a leitura do numero do usuario
        bNumero = Console.ReadLine()
        sNumeroExtenso = objConversor.Numeros0e19(bNumero)

        objConversor.validarCasasDecimais(Convert.ToString(bNumero))

        Console.WriteLine("Valor por Extenso: " + sNumeroExtenso)

        Console.WriteLine("")
        Console.WriteLine("")
        Console.WriteLine("Group Software - 2018 ")
        bNumero = Console.ReadLine()
    End Sub



End Module
