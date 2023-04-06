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
'Definindo variáveis
Dim v1 As Integer
Dim v2 As Integer
Dim v3 As Integer
Dim v4 As Integer
Dim result As Long

'Definido qual planilha será executada no código
Dim w As Worksheet

'Ajustando, nome igual a planilha, com isso W sendo definido como plan1
Set w = Sheets("Plan1")

w.Select 'Apenas para selecionar a planilha


v1 = w.Range("a1")
v2 = w.Range("a2")
v3 = w.Range("a3")
v4 = w.Range("a4")

result = v1 + v2
MsgBox result

result = v3 * v4
MsgBox result

MsgBox v1 > v2

End Sub
