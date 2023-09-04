Attribute VB_Name = "Aula13_Private_e_Publico"
Private textoPrivado As String
Public textoPublico As String

Sub valorPrivado()

    textoPrivado = "VBA Modo Privado"

End Sub

Sub valorPublico()

    textoPublico = "VBA modo Publico "

End Sub


Sub Main()
    Dim nome As String
    Dim sobrenome As String
    
    nome = "Jonathan"
    sobrenome = "Nogueira"
    
    
    Call valorPublico
    Call valorPrivado
    
    MsgBox "Nome: " & nome & vbCrLf & _
            "Sobrenome: " & sobrenome & vbCrLf & _
            "Texto publico: " & textoPublico & vbCrLf & _
            "Texto Privado:  " & textoPrivado
    

End Sub

