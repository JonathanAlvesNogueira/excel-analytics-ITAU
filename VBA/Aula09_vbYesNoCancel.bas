Attribute VB_Name = "Aula09_vbYesNoCancel"
Sub YesNoCancel()

   resultado = MsgBox("Quer continuar ? ", vbYesNoCancel)
   If (resultado = vbYes) Then
        MsgBox "Deletado", Title:="Executado"
    ElseIf (resultado = vbNo) Then
        MsgBox "Operação Cancelada", Title:="Não executado"
    Else
        MsgBox ("Ação não executada")
    End If
    
End Sub

