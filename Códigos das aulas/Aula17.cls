VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub bt_executar1_Click()
'Declarando variáveis
Dim w           As Worksheet
Dim resultado   As Long

'Inicializando variável
Set w = Sheets("Plan1")
resultado = 0

'Selecionando planilha
w.Select
w.Range("a2").Select

'Estrutura de repetição
Do Until resultado >= 4000
    resultado = resultado + ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
Loop

'Estrutura de repetição para chegar na última célula
Do While ActiveCell.Value <> ""
    ActiveCell.Offset(1, 0).Select
Loop

'Colocando o resultado na célula final
ActiveCell.Value = "Célula final" & resultado

End Sub
