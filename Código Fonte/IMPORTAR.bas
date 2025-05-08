Attribute VB_Name = "M�dulo1"
Sub AtualizarPlanilhasDeOutroArquivo()
    ' Desativar a atualiza��o de tela para melhorar a performance
    On Error GoTo RestaurarConfiguracoes
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Desativar filtros na planilha de destino de forma robusta
    On Error Resume Next
    With ThisWorkbook.Sheets("Resultados")
        If .AutoFilterMode Then
            If .FilterMode Then .ShowAllData
        End If
    End With
    On Error GoTo 0

    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wbDestino As Workbook
    Dim wsEstoque As Worksheet
    Dim wsReversa As Worksheet
    Dim ultimaLinhaOrigem As Long
    Dim ultimaLinhaEstoque As Long
    Dim ultimaLinhaReversa As Long
    Dim caminhoOrigem As String
    Dim caminhoDestino As String
    Dim i As Long, j As Long
    Dim encontrado As Boolean
    Dim serialNovoDestino As String
    Dim serialNovoOrigem As String
    Dim serialRetirado As String
    Dim modeloCorrespondente As String
    Dim duplicados As Object
    Dim contadorEstoque As Long ' Vari�vel para contar os registros atualizados na planilha ESTOQUE
    Dim contadorReversa As Long ' Vari�vel para contar os registros atualizados na planilha REVERSA

    ' Inicializar contadores de atualiza��es
    contadorEstoque = 0
    contadorReversa = 0

    ' Obter o caminho do arquivo CSV da c�lula B1 da planilha Importar (arquivo IMPORTAR)
    caminhoOrigem = ThisWorkbook.Sheets("Importar").Range("B1").Value

    If caminhoOrigem = "" Then
        MsgBox "Erro: Nenhum arquivo CSV foi selecionado.", vbCritical
        GoTo Limpeza
    End If

    ' Abrir o arquivo CSV como uma pasta de trabalho
    On Error Resume Next
    Set wbOrigem = Workbooks.Open(caminhoOrigem, Local:=True)
    On Error GoTo 0

    If wbOrigem Is Nothing Then
        MsgBox "Erro: O arquivo CSV '" & caminhoOrigem & "' n�o foi encontrado ou n�o p�de ser aberto.", vbCritical
        GoTo Limpeza
    End If

    ' Obter o caminho do arquivo ESTOQUE
    caminhoDestino = ThisWorkbook.Path & "\ESTOQUE.xlsm"

    ' Abrir o arquivo ESTOQUE como uma pasta de trabalho
    On Error Resume Next
    Set wbDestino = Workbooks.Open(caminhoDestino, ReadOnly:=False)
    On Error GoTo 0

    If wbDestino Is Nothing Then
        MsgBox "Erro: O arquivo ESTOQUE n�o foi encontrado no mesmo diret�rio.", vbCritical
        wbOrigem.Close False
        GoTo Limpeza
    End If

    ' Configura��o das planilhas
    Set wsOrigem = wbOrigem.Sheets(1) ' Primeira planilha do CSV
    Set wsEstoque = wbDestino.Sheets("ESTOQUE") ' Planilha ESTOQUE no arquivo ESTOQUE
    Set wsReversa = wbDestino.Sheets("REVERSA") ' Planilha REVERSA no arquivo ESTOQUE

    If wsOrigem Is Nothing Or wsEstoque Is Nothing Or wsReversa Is Nothing Then
        MsgBox "Erro: As planilhas n�o foram configuradas corretamente.", vbCritical
        wbOrigem.Close False
        wbDestino.Close False
        GoTo Limpeza
    End If

    ' Obter �ltimas linhas
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
    ultimaLinhaEstoque = wsEstoque.Cells(wsEstoque.Rows.Count, "E").End(xlUp).Row
    ultimaLinhaReversa = wsReversa.Cells(wsReversa.Rows.Count, "D").End(xlUp).Row + 1

    ' Inicializar dicion�rio para evitar duplicados
    Set duplicados = CreateObject("Scripting.Dictionary")

    ' Preencher dicion�rio com seriais existentes na planilha REVERSA
    For i = 2 To ultimaLinhaReversa - 1
        duplicados(Trim(UCase(wsReversa.Cells(i, "D").Value))) = True
    Next i

    ' Atualizar a planilha ESTOQUE
    For i = 2 To ultimaLinhaEstoque
        ' Verificar se o status � "Ativado", e se for, pular a atualiza��o
        If UCase(wsEstoque.Cells(i, 1).Value) = "ATIVADO" Then
            GoTo ProximaLinhaEstoque ' Pular para a pr�xima linha se j� estiver ativado
        End If

        serialNovoDestino = Trim(UCase(wsEstoque.Cells(i, 5).Value)) ' Serial novo da planilha ESTOQUE em mai�sculas
        encontrado = False

        If serialNovoDestino <> "" Then
            For j = 2 To ultimaLinhaOrigem
                serialNovoOrigem = Trim(UCase(wsOrigem.Cells(j, 30).Value)) ' Serial novo da planilha origem em mai�sculas
                Dim statusAtual As String
                statusAtual = UCase(wsOrigem.Cells(j, 3).Value)

                ' Verificar se o serial corresponde e se o status � FINALIZADO
                If serialNovoDestino = serialNovoOrigem Then
                    If statusAtual = "FINALIZADO" Then
                        ' Preencher informa��es na planilha ESTOQUE
                        wsEstoque.Cells(i, 3).Value = wsOrigem.Cells(j, 23).Value ' T�cnico (Coluna W)
                        wsEstoque.Cells(i, 6).Value = wsOrigem.Cells(j, 19).Value ' Data de Atendimento (Coluna S)
                        wsEstoque.Cells(i, 7).Value = wsOrigem.Cells(j, 1).Value  ' Ordem de Servi�o (Coluna A)
                        wsEstoque.Cells(i, 8).Value = UCase(wsOrigem.Cells(j, 29).Value) ' Serial Equipamento Antigo (Coluna AC)
                        wsEstoque.Cells(i, 1).Value = "Ativado" ' Atualizar status para "Ativado"
                        encontrado = True
                        contadorEstoque = contadorEstoque + 1 ' Incrementar contador de atualiza��es na planilha ESTOQUE
                        Exit For
                    End If
                End If
            Next j

            If Not encontrado Then
                wsEstoque.Cells(i, 1).Value = "Base" ' Manter status como "Base" se n�o encontrado
            End If

            ' Formatar a data de atendimento para DD/MM/AAAA
            Dim valorCelula As Variant

            valorCelula = wsEstoque.Cells(i, 6).Value

            If IsDate(valorCelula) Then
                wsEstoque.Cells(i, 6).Value = Format(CDate(valorCelula), "DD/MM/YYYY")
            End If

        End If
ProximaLinhaEstoque:
    Next i

    ' Atualizar a planilha REVERSA
    For i = 2 To ultimaLinhaEstoque
        serialRetirado = Trim(UCase(wsEstoque.Cells(i, "H").Value)) ' Serial retirado em mai�sculas
        modeloCorrespondente = Trim(wsEstoque.Cells(i, "D").Value) ' Modelo correspondente

        If serialRetirado <> "" Then
            ' Verificar se o serial j� est� na planilha REVERSA para evitar duplicatas
            If Not duplicados.Exists(serialRetirado) Then
                duplicados.Add serialRetirado, True
                wsReversa.Cells(ultimaLinhaReversa, "D").Value = serialRetirado ' Serial retirado na coluna D
                wsReversa.Cells(ultimaLinhaReversa, "B").Value = "BAD" ' Status BAD na coluna B
                wsReversa.Cells(ultimaLinhaReversa, "C").Value = modeloCorrespondente ' Modelo na coluna C
                ultimaLinhaReversa = ultimaLinhaReversa + 1
                contadorReversa = contadorReversa + 1 ' Incrementar contador de atualiza��es na planilha REVERSA
            End If
        End If
    Next i

    ' Fechar o arquivo de origem
    wbOrigem.Close False

    ' Salvar e fechar o arquivo de destino, somente se houver mudan�as
    If contadorEstoque > 0 Or contadorReversa > 0 Then
        wbDestino.Save
    End If

    ' Exibir mensagem com a quantidade de registros atualizados em cada planilha
    MsgBox "Atualiza��o Conclu�da com Sucesso! " & contadorEstoque & " Novos Registros no ESTOQUE." & vbCrLf & _
           contadorReversa & " Novos Registros na REVERSA.", vbInformation

Limpeza:
    ' Restaurar configura��es
RestaurarConfiguracoes:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Erro: " & Err.Description, vbCritical
    End If

End Sub





Sub SelecionarArquivo()
    Dim dialogoArquivo As FileDialog
    Dim caminho As String

    ' Abrir di�logo para selecionar o arquivo
    Set dialogoArquivo = Application.FileDialog(msoFileDialogFilePicker)
    dialogoArquivo.Title = "Selecione o arquivo CSV"
    dialogoArquivo.Filters.Clear
    dialogoArquivo.Filters.Add "Arquivos CSV", "*.csv", 1

    If dialogoArquivo.Show = -1 Then
        caminho = dialogoArquivo.SelectedItems(1)
        ThisWorkbook.Sheets("Importar").Range("B1").Value = caminho
    End If
End Sub








