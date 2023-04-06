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
'Soma de valores da planilha
'Dimensão de variáveis
Dim resultado   As Long
Dim w           As Worksheet

'definir planilha
Set w = Sheets("Plan1")
resultado = 0 'Zerando o valor para mais de uma execução

w.Select 'Selecionando a planilha
w.Range("a1").Select 'Selecionando a célula da planilha

'Estrutura de repetição
Do While ActiveCell.Value <> ""
'Colocar a ação de repetição
    'Resultado
    '------------------------
    
    resultado = resultado + ActiveCell.Value 'somando os valores de células ativas
    ActiveCell.Offset(1, 0).Select 'Deslocando a célula ativa para baixo

Loop

'Saída de somatório
MsgBox "O somatório é : " & resultado

End Sub
