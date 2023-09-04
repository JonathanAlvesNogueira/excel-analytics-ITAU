Attribute VB_Name = "Aula05_Concatenando_mensagem"
Sub concatenaMensage()

Dim celula1 As String
Dim celula2 As String

celula1 = Range("A1")
celula2 = Range("B2")

' CONCATENA VALOR DE DUAS CELULAS
MsgBox (celula1 & " e " & celula2)

End Sub

