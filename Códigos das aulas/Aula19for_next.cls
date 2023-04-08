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

'Definindo variáveis
Dim w           As Worksheet
Dim count       As Integer

'Inicializando
Set w = Sheets("Plan1")

'selecionando
w.Select
w.Range("a1").Select

'Estrutura de repetição com inicio e parada
For count = 1 To 20 'repetição começando de 1 indo até 20 contador fixo
'Colocar a repetição do procedimento
    ActiveCell.Value = count
    ActiveCell.Offset(1, 0).Select
Next
End Sub
