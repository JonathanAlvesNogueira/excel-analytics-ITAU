Attribute VB_Name = "Aula12_private_VBA"
Private texto As String

Sub atribuiValor()

    texto = "Estuda VBA"

End Sub

Sub escreveValor()

    Dim nome As String
    nome = "Jonathan"
    
    ' CHAMA O METODO ATRIBUIVALOR()
    Call atribuiValor
    
    MsgBox "Nome: " & nome & "  " & texto

End Sub
