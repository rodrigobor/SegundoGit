Attribute VB_Name = "Módulo1"
Sub estrutura()
'para declarar variavel no VBA usamos o comando DIM
Dim produto As String
Dim preco As Double
Dim desconto As Double
Dim precofinal As Double

'vamos utilizar a caixa de entrada inputbox para as variáveis
produto = InputBox("digite o nome do produto", "produto")
preco = InputBox("digite o preço do produto", "preço")
desconto = InputBox("digite o desconto", "desconto")
precofinal = preco - preco * desconto

Range("A1").Value = prodruto
Range("A2").Value = preco
Range("A3").Value = desconto
Range("A4").Value = precofinal

End Sub
