Attribute VB_Name = "Aula09_vbYesNoCancel"
Sub YesNoCancel()

   resultado = MsgBox("Quer continuar ? ", vbYesNoCancel)
   If (resultado = vbYes) Then
        MsgBox "Deletado", Title:="Executado"
    ElseIf (resultado = vbNo) Then
        MsgBox "Opera��o Cancelada", Title:="N�o executado"
    Else
        MsgBox ("A��o n�o executada")
    End If
    
End Sub

