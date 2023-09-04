Attribute VB_Name = "Aula10_String_e_inputBox"
Sub variavelStringEByte()

'STRING = " TEXTO "
Dim nome As String
Dim sobrenome As String

nome = InputBox("Informe o nome", Title:=Cadastro)

sobrenome = InputBox("Informe o seu sobrenome", Title:=Cadastro)


MsgBox "Nome: " & nome & " Sobrenome: " & sobrenome


End Sub
