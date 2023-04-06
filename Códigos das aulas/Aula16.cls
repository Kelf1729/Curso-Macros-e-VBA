VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

'Declarando variáveis
Dim w           As Worksheet
Dim ultcell     As Range
Dim resultado   As Long


'inicializando elas
Set w = Sheets("Plan1")

'última célula
Set ultcell = w.Range("a1048576").End(xlUp)
ultcell.Select

'Subindo as células
Do While ActiveCell.Row >= 2
    resultado = resultado + ActiveCell.Value 'Somando
    ActiveCell.Offset(-1, 0).Select 'Subindo as linhas
Loop

'Resultado
MsgBox "A soma dos valores é :" & resultado

End Sub
