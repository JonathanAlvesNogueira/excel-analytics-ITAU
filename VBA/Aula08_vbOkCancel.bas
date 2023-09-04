Attribute VB_Name = "Aula08_vbOkCancel"
Sub mensagemTipoOkCancel()

' GERA UMA MENSAGEMBOX DE OK OU CANCELA
MsgBox "Relatório gerado com sucesso! ", vbOKCancel

End Sub

Sub mensagemTipoOkCancel_2()

' GERA UMA MENSAGEMBOX DE OK OU CANCELA
'SE O CLICK FOR IGUAL OK IMPRIME O FEITO COM SUCESSO SE NAO DA ERRO
If (MsgBox("Relatório gerado com sucesso! ", vbOKCancel, Title = "Tem certeza")) = vbOK Then
    MsgBox "Feito com sucesso", Title:="Realizado"
Else
    MsgBox "Cancelado com sucesso", vbCritical, "Alerta!!!!"
End If


End Sub

