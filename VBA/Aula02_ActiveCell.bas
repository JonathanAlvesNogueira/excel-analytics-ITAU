Attribute VB_Name = "Aula02_ActiveCell"
' OBRIGA QUE TODAS AS VARIÁVEIS TENHAM QUE SER DECLARADAS SEU TIPO
Option Explicit
Sub Cellp()

Dim numero As Integer

numero = 10
'  COLOCA O VALOR 10 NA CÉLULA QUE ESTÁ SELECIONADA NO EXCEL

ActiveCell.Value = numero


End Sub
