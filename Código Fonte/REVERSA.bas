Attribute VB_Name = "Módulo1"
Option Explicit

Sub ChecarSerialPorBusca()
    Dim wbEstoque As Workbook
    Dim wsReversa As Worksheet
    Dim wsRecebimento As Worksheet
    Dim ultimaLinhaReversa As Long
    Dim serialBusca As String
    Dim encontrado As Boolean
    Dim i As Long
    Dim celulaChecada As Range
    Dim shape As shape
    Dim textBox As Object
    Dim caminhoEstoque As String

    ' Caminho do arquivo ESTOQUE
    caminhoEstoque = ThisWorkbook.Path & "\ESTOQUE.xlsm"
    
    ' Abrir o arquivo ESTOQUE se não estiver aberto
    On Error Resume Next
    Set wbEstoque = Workbooks("ESTOQUE.xlsm")
    On Error GoTo 0
    
    If wbEstoque Is Nothing Then
        On Error Resume Next
        Set wbEstoque = Workbooks.Open(caminhoEstoque, ReadOnly:=True)
        On Error GoTo 0
        If wbEstoque Is Nothing Then
            MsgBox "Erro: O arquivo ESTOQUE não foi encontrado no mesmo diretório.", vbCritical
            Exit Sub
        End If
    End If

    ' Definir referências às planilhas
    On Error Resume Next
    Set wsReversa = wbEstoque.Sheets("REVERSA") ' Planilha REVERSA no arquivo ESTOQUE
    Set wsRecebimento = ThisWorkbook.Sheets("RECEBIMENTO") ' Planilha RECEBIMENTO no arquivo REVERSAS
    On Error GoTo 0

    ' Verificar se as planilhas foram definidas corretamente
    If wsReversa Is Nothing Then
        MsgBox "A planilha 'REVERSA' não foi encontrada no arquivo ESTOQUE.", vbCritical
        If Not wbEstoque Is Nothing Then wbEstoque.Close False
        Exit Sub
    End If

    If wsRecebimento Is Nothing Then
        MsgBox "A planilha 'RECEBIMENTO' não foi encontrada no arquivo REVERSAS.", vbCritical
        Exit Sub
    End If

    ' Obter o serial da caixa de texto
    Set textBox = wsRecebimento.OLEObjects("TextBox1").Object
    serialBusca = UCase(Trim(textBox.Text)) ' Converter para maiúsculas e remover espaços extras

    ' Verificar se o serial foi preenchido
    If serialBusca = "" Then
        MsgBox "Por favor, insira ou escaneie um serial na caixa de texto.", vbExclamation
        Exit Sub
    End If

    ' Obter última linha preenchida na planilha REVERSA
    ultimaLinhaReversa = wsReversa.Cells(wsReversa.Rows.Count, "D").End(xlUp).Row

    ' Procurar o serial na planilha REVERSA
    encontrado = False
    For i = 2 To ultimaLinhaReversa
        If UCase(Trim(wsReversa.Cells(i, "D").Value)) = serialBusca Then ' Comparar ignorando maiúsculas/minúsculas
            ' Definir a célula onde será inserido o tique
            Set celulaChecada = wsReversa.Cells(i, "E")

            ' Remover qualquer texto existente na célula
            celulaChecada.ClearContents

            ' Inserir um tique como texto na célula
            celulaChecada.Font.Name = "Wingdings"
            celulaChecada.Font.Size = 14
            celulaChecada.Font.Color = RGB(0, 176, 80) ' Cor verde
            celulaChecada.Value = Chr(252) ' Código do tique em Wingdings

            encontrado = True
            Exit For
        End If
    Next i

    ' Remover formas existentes na planilha RECEBIMENTO
    For Each shape In wsRecebimento.Shapes
        If shape.Name Like "Resultado*" Then shape.Delete
    Next shape

    ' Se encontrado, adicionar tique verde (dobrando o tamanho)
If encontrado Then
    Set shape = wsRecebimento.Shapes.AddShape(msoShapeOval, 100, 50, 120, 120) ' Dobrar o tamanho
    shape.Fill.ForeColor.RGB = RGB(0, 176, 80) ' Cor verde
    shape.Line.ForeColor.RGB = RGB(0, 128, 0) ' Cor verde escuro
    shape.Line.Weight = 2
    shape.TextFrame2.TextRange.Text = ChrW(&H2713) ' Símbolo de tique
    shape.TextFrame2.TextRange.Font.Size = 64 ' Aumentar o tamanho do texto
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' Cor branca para o tique
    shape.TextFrame2.HorizontalAnchor = msoAnchorCenter
    shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shape.Name = "ResultadoTique"
Else
    ' Se não encontrado, adicionar X vermelho (dobrando o tamanho)
    Set shape = wsRecebimento.Shapes.AddShape(msoShapeOval, 100, 50, 120, 120) ' Dobrar o tamanho
    shape.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Cor vermelha
    shape.Line.ForeColor.RGB = RGB(128, 0, 0) ' Cor vermelha escura
    shape.Line.Weight = 2
    shape.TextFrame2.TextRange.Text = "X"
    shape.TextFrame2.TextRange.Font.Size = 64 ' Aumentar o tamanho do texto
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' Cor branca para o X
    shape.TextFrame2.HorizontalAnchor = msoAnchorCenter
    shape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shape.Name = "ResultadoX"
End If

    ' Retornar mensagem de status
    If encontrado Then
        MsgBox "Serial " & serialBusca & " encontrado e marcado com um tique na planilha REVERSA.", vbInformation
    Else
        MsgBox "Serial " & serialBusca & " não encontrado na planilha REVERSA.", vbExclamation
    End If

    ' Limpar a caixa de texto
    textBox.Text = ""

    
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' 13 é o código da tecla Enter
        Call ChecarSerialPorBusca
    End If
End Sub



