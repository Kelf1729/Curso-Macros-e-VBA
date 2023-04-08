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
'Definir variáveis
Dim w           As Worksheet
Dim resultado   As Long

'Inicializando variáveis
Set w = Sheets("Plan1")
resultado = 0

'Selecionar elementos
w.Select
w.Range("a2").Select 'Selecionando a célula A2

'Estrutura de loop para percorrer células
Do While ActiveCell.Value <> ""
    'Somando apenas valores positivos com if then
        If ActiveCell.Value > 0 Then'Estrutura condicional if ... then 
        resultado = resultado + ActiveCell.Value
        ActiveCell.Offset(0, 1).Value = 1 'Veriricar os positivos
    End If
    ActiveCell.Offset(1, 0).Select 'Desce para a próxima linha
Loop

MsgBox resultado

End Sub
