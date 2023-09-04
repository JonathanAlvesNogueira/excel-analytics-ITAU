Attribute VB_Name = "Aula11_replace"
Sub mudandoNome()

Dim nomeCompleto As String

nomeCompleto = InputBox("Informe seu nome")
nomeCompleto = Replace(nomeCompleto, "Alves", "Adalberto")

MsgBox ("Essa é a Mudança final " & nomeCompleto)


End Sub
