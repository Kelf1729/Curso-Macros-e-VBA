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
Dim count           As Integer

'Declarando número de planilhas
count = Sheets.count

For Each planilhas In Sheets 'Para cada planilha das planilhas
'O próprio VBA cria a variável Planilhas
    planilhas.Protect Password:="123" 'Criando uma proteção em cada planilha com senha 123
Next
End Sub

Private Sub bt_executar2_Click()
For Each planilha In Sheets
    'tirando proteção das planilhas
    planilha.Unprotect Password:="123"
Next
End Sub
