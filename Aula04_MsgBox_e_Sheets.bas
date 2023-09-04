Attribute VB_Name = "Aula04_MsgBox_e_Sheets"
Sub mensagemRangeSheets()

Dim celula As String
celula = Range("B10")
' APONTA PARA A CELULA DA PLANILHA PROPOSTA
MsgBox Sheets("proposta").Range("B10")
End Sub

