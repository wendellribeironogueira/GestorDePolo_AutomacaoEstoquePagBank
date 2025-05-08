Attribute VB_Name = "Módulo1"
Sub GerarPainelDeInformacoes()
    Dim ws As Worksheet
    Dim tempWs As Worksheet
    Dim csvFile As String
    Dim qt As QueryTable
    Dim lastRow As Long
    Dim slaRange As Range
    Dim slaCell As Range
    Dim resultadosWs As Worksheet
    Dim graficosWs As Worksheet
    Dim dictTecnicos As Object
    Dim dictStatus As Object
    Dim dictTipoServico As Object
    Dim dictMotivoCancelamento As Object
    Dim dictChamadosNovos As Object
    Dim dictChamadosFinalizados As Object
    Dim dictChamadosImprodutivos As Object
    Dim totalChamadosNovos As Long
    Dim totalChamadosFinalizados As Long
    Dim totalChamadosImprodutivos As Long
    Dim dentroSLA As Long
    Dim foraSLA As Long
    Dim i As Long
    Dim currentTime As Date
    Dim dictChamadosNovosPorServico As Object
    Dim dictChamadosForaHorario As Object
    Dim totalForaHorario As Long
    Dim periodo As String
    Dim dictTentativasReabertura As Object
    Dim dictAlertasSLA As Object
    Dim dictChamadosImprodutivosDetalhados As Object
    Dim horasRestantes As Double
    Dim horas As Long
    Dim minutos As Long
    Dim minDate As Date
    Dim maxDate As Date
    Dim dictChamadosVencendoHoje As Object
    Dim dictChamadosFinalizadosMes As Object
    Dim dictTecnicosDetalhados As Object
    Dim dictChamadosProcessados As Object
    Dim dictModelos As Object
    ' Obter a data e hora atuais
    currentTime = Now

    ' Selecionar o arquivo CSV
    csvFile = Application.GetOpenFilename("Arquivos CSV (*.csv), *.csv", , "Selecione o arquivo CSV")
    If csvFile = "False" Then Exit Sub

    ' Importar o arquivo CSV para uma planilha temporária
    Set tempWs = Worksheets.Add
    Set qt = tempWs.QueryTables.Add(Connection:="TEXT;" & csvFile, Destination:=tempWs.Range("A1"))
    With qt
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFilePlatform = xlWindows
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .Refresh BackgroundQuery:=False
    End With

    ' Converter colunas de data para o formato datetime
    tempWs.Columns("P:Q").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    tempWs.Columns("S:S").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    ' Calcular o SLA de 24 horas a partir da data de abertura
    lastRow = tempWs.Cells(tempWs.Rows.Count, "A").End(xlUp).Row
    tempWs.Cells(1, "AA").Value = "SLA"
    dentroSLA = 0
    foraSLA = 0
    Set slaRange = tempWs.Range("AA2:AA" & lastRow)

    ' Inicializar dicionários para armazenar os dados
    Set dictTecnicos = CreateObject("Scripting.Dictionary")
    Set dictStatus = CreateObject("Scripting.Dictionary")
    Set dictTipoServico = CreateObject("Scripting.Dictionary")
    Set dictMotivoCancelamento = CreateObject("Scripting.Dictionary")
    Set dictChamadosNovos = CreateObject("Scripting.Dictionary")
    Set dictChamadosFinalizados = CreateObject("Scripting.Dictionary")
    Set dictChamadosImprodutivos = CreateObject("Scripting.Dictionary")
    Set dictChamadosNovosPorServico = CreateObject("Scripting.Dictionary")
    Set dictChamadosForaHorario = CreateObject("Scripting.Dictionary")
    Set dictTentativasReabertura = CreateObject("Scripting.Dictionary")
    Set dictAlertasSLA = CreateObject("Scripting.Dictionary")
    Set dictChamadosImprodutivosDetalhados = CreateObject("Scripting.Dictionary")
    Set dictChamadosVencendoHoje = CreateObject("Scripting.Dictionary")
    Set dictChamadosFinalizadosMes = CreateObject("Scripting.Dictionary")
    Set dictTecnicosDetalhados = CreateObject("Scripting.Dictionary")
    Set dictChamadosProcessados = CreateObject("Scripting.Dictionary")
    Set dictModelos = CreateObject("Scripting.Dictionary")

    totalChamadosNovos = 0
    totalChamadosFinalizados = 0
    totalChamadosImprodutivos = 0
    totalForaHorario = 0

    ' Ajuste da lógica de SLA
    For Each slaCell In slaRange
        If IsDate(tempWs.Cells(slaCell.Row, "P").Value) And IsDate(tempWs.Cells(slaCell.Row, "Q").Value) And IsDate(tempWs.Cells(slaCell.Row, "S").Value) Then
            Dim dataAbertura As Date
            Dim dataLimite As Date
            Dim dataFechamento As Date
            dataAbertura = CDate(tempWs.Cells(slaCell.Row, "P").Value)
            dataLimite = CDate(tempWs.Cells(slaCell.Row, "Q").Value)
            dataFechamento = CDate(tempWs.Cells(slaCell.Row, "S").Value)
            horasRestantes = (dataLimite - Now) * 24

            If dataFechamento <= dataLimite Then
                slaCell.Value = "Dentro do SLA"
                dentroSLA = dentroSLA + 1
            Else
                slaCell.Value = "Fora do SLA"
                foraSLA = foraSLA + 1
            End If

            ' Adicionar alerta com número do chamado e tempo restante para o SLA expirar, apenas para chamados "Encaminhado"
            If tempWs.Cells(slaCell.Row, "C").Value = "Encaminhado" Then
                horas = Int(horasRestantes)
                minutos = (horasRestantes - horas) * 60
                dictAlertasSLA.Add tempWs.Cells(slaCell.Row, "A").Value, "Faltam " & horas & " horas e " & Format(minutos, "00") & " minutos para Fim do SLA"
            End If
        Else
            slaCell.Value = "Data inválida"
        End If
    Next slaCell

    ' Coletar dados
    minDate = tempWs.Cells(2, "P").Value
    maxDate = tempWs.Cells(2, "P").Value
    For i = 2 To lastRow
        ' Verificar se o chamado já foi processado
        If Not dictChamadosProcessados.exists(tempWs.Cells(i, "A").Value) Then
            dictChamadosProcessados.Add tempWs.Cells(i, "A").Value, True

            ' Atualizar minDate e maxDate
            If IsDate(tempWs.Cells(i, "P").Value) Then
                If tempWs.Cells(i, "P").Value < minDate Then minDate = tempWs.Cells(i, "P").Value
                If tempWs.Cells(i, "P").Value > maxDate Then maxDate = tempWs.Cells(i, "P").Value
            End If

            ' Coletar dados por Chamado Finalizado por Técnico
            If tempWs.Cells(i, "C").Value = "Finalizado" Then
                If Not dictTecnicos.exists(tempWs.Cells(i, "W").Value) Then
                    dictTecnicos(tempWs.Cells(i, "W").Value) = 0
                End If
                dictTecnicos(tempWs.Cells(i, "W").Value) = dictTecnicos(tempWs.Cells(i, "W").Value) + 1

                ' Coletar dados por Chamado Finalizado por Serviço
                If Not dictChamadosFinalizados.exists(tempWs.Cells(i, "N").Value) Then
                    dictChamadosFinalizados(tempWs.Cells(i, "N").Value) = 0
                End If
                dictChamadosFinalizados(tempWs.Cells(i, "N").Value) = dictChamadosFinalizados(tempWs.Cells(i, "N").Value) + 1
                totalChamadosFinalizados = totalChamadosFinalizados + 1

                ' Acumulativo do mês de chamados finalizados dentro e fora do prazo
                If Month(dataFechamento) = Month(currentTime) And Year(dataFechamento) = Year(currentTime) Then
                    If Not dictChamadosFinalizadosMes.exists("Dentro do Prazo") Then
                        dictChamadosFinalizadosMes("Dentro do Prazo") = 0
                    End If
                    If Not dictChamadosFinalizadosMes.exists("Fora do Prazo") Then
                        dictChamadosFinalizadosMes("Fora do Prazo") = 0
                    End If
                    If dataFechamento <= dataLimite Then
                        dictChamadosFinalizadosMes("Dentro do Prazo") = dictChamadosFinalizadosMes("Dentro do Prazo") + 1
                    Else
                        dictChamadosFinalizadosMes("Fora do Prazo") = dictChamadosFinalizadosMes("Fora do Prazo") + 1
                    End If
                End If
            End If
             ' Coletar dados por status
            If Not dictStatus.exists(tempWs.Cells(i, "C").Value) Then
                dictStatus(tempWs.Cells(i, "C").Value) = 0
            End If
            dictStatus(tempWs.Cells(i, "C").Value) = dictStatus(tempWs.Cells(i, "C").Value) + 1

            ' Coletar dados de chamados finalizados, improdutivos e encaminhados
            If tempWs.Cells(i, "C").Value = "Finalizado" Or tempWs.Cells(i, "C").Value = "Improdutivo" Or tempWs.Cells(i, "C").Value = "Encaminhado" Then
                If Not dictChamadosNovos.exists(tempWs.Cells(i, "N").Value) Then
                    dictChamadosNovos(tempWs.Cells(i, "N").Value) = 0
                End If
                dictChamadosNovos(tempWs.Cells(i, "N").Value) = dictChamadosNovos(tempWs.Cells(i, "N").Value) + 1
                totalChamadosNovos = totalChamadosNovos + 1

                If Not dictChamadosNovosPorServico.exists(tempWs.Cells(i, "N").Value) Then
                    dictChamadosNovosPorServico(tempWs.Cells(i, "N").Value) = 0
                End If
                dictChamadosNovosPorServico(tempWs.Cells(i, "N").Value) = dictChamadosNovosPorServico(tempWs.Cells(i, "N").Value) + 1
            End If

            If tempWs.Cells(i, "C").Value = "Encaminhado" Then
                ' Calcular o tempo restante para o SLA expirar
                If IsDate(tempWs.Cells(i, "Q").Value) Then
                    horasRestantes = (CDate(tempWs.Cells(i, "Q").Value) - Now) * 24
                    horas = Int(horasRestantes)
                    minutos = (horasRestantes - horas) * 60
                    dictAlertasSLA.Add tempWs.Cells(i, "A").Value, "Faltam " & horas & " horas e " & Format(minutos, "00") & " minutos para Fim do SLA"
                End If
            End If

            If tempWs.Cells(i, "C").Value = "Improdutivo" Then
                If Not dictChamadosImprodutivos.exists(tempWs.Cells(i, "N").Value) Then
                    dictChamadosImprodutivos(tempWs.Cells(i, "N").Value) = 0
                End If
                dictChamadosImprodutivos(tempWs.Cells(i, "N").Value) = dictChamadosImprodutivos(tempWs.Cells(i, "N").Value) + 1
                totalChamadosImprodutivos = totalChamadosImprodutivos + 1
            End If

            ' Coletar dados de chamados fora do horário comercial
            abertura = tempWs.Cells(i, "P").Value
            If Hour(abertura) >= 17 Or Hour(abertura) < 9 Then
                If Not dictChamadosForaHorario.exists(tempWs.Cells(i, "N").Value) Then
                    dictChamadosForaHorario(tempWs.Cells(i, "N").Value) = 0
                End If
                dictChamadosForaHorario(tempWs.Cells(i, "N").Value) = dictChamadosForaHorario(tempWs.Cells(i, "N").Value) + 1
                totalForaHorario = totalForaHorario + 1
            End If

            ' Verificar se o chamado está "Improdutivo" por "Cliente Ausente", "Estabelecimento Fechado" ou "Endereço Incorreto" e reabri-lo
            If tempWs.Cells(i, "C").Value = "Improdutivo" And _
               (tempWs.Cells(i, "AU").Value = "Cliente Ausente" Or _
                tempWs.Cells(i, "AU").Value = "Estabelecimento Fechado" Or _
                tempWs.Cells(i, "AU").Value = "Endereço Incorreto") Then
                tempWs.Cells(i, "C").Value = "Reaberto"
                If Not dictTentativasReabertura.exists(tempWs.Cells(i, "A").Value) Then
                    dictTentativasReabertura(tempWs.Cells(i, "A").Value) = 0
                End If
                dictTentativasReabertura(tempWs.Cells(i, "A").Value) = dictTentativasReabertura(tempWs.Cells(i, "A").Value) + 1

                ' Verificar se o chamado já foi reaberto duas vezes
                If dictTentativasReabertura(tempWs.Cells(i, "A").Value) >= 2 Then
                    MsgBox "Chamado " & tempWs.Cells(i, "A").Value & " foi reaberto duas vezes e será finalizado em definitivo.", vbInformation
                    tempWs.Cells(i, "C").Value = "Finalizado"
                End If
            End If

            ' Coletar dados de chamados vencendo hoje
            If tempWs.Cells(i, "C").Value = "Encaminhado" And DateValue(tempWs.Cells(i, "Q").Value) = DateValue(currentTime) Then
                If Not dictChamadosVencendoHoje.exists(tempWs.Cells(i, "A").Value) Then
                    dictChamadosVencendoHoje(tempWs.Cells(i, "A").Value) = tempWs.Cells(i, "Q").Value
                End If
            End If
' Condição para verificar o status do chamado
        If tempWs.Cells(i, "C").Value = "ENCAMINHADO" Or tempWs.Cells(i, "C").Value = "NOVA" Then
            ' Extrair o modelo do equipamento da coluna AA
            modelo = tempWs.Cells(i, "AA").Value

            ' Atualizar a contagem no dicionário
            If Not dictModelos.exists(modelo) Then
                dictModelos(modelo) = 0
            End If
            dictModelos(modelo) = dictModelos(modelo) + 1
        End If
        
            ' Coletar dados detalhados por técnico
            Dim tecnico As String
            tecnico = tempWs.Cells(i, "W").Value
            If Not dictTecnicosDetalhados.exists(tecnico) Then
                Set dictTecnicosDetalhados(tecnico) = CreateObject("Scripting.Dictionary")
                dictTecnicosDetalhados(tecnico).Add "Improdutivo", 0
                dictTecnicosDetalhados(tecnico).Add "Finalizado", 0
                dictTecnicosDetalhados(tecnico).Add "Encaminhado", 0
            End If
            dictTecnicosDetalhados(tecnico)(tempWs.Cells(i, "C").Value) = dictTecnicosDetalhados(tecnico)(tempWs.Cells(i, "C").Value) + 1
        End If
    Next i

    ' Coletar dados por cidade
    Dim dictChamadosPorCidade As Object
    Set dictChamadosPorCidade = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        Dim cidade As String
        cidade = tempWs.Cells(i, "G").Value

        If Not dictChamadosPorCidade.exists(cidade) Then
            Set dictChamadosPorCidade(cidade) = CreateObject("Scripting.Dictionary")
        End If

        Dim tipoServico As String
        tipoServico = tempWs.Cells(i, "N").Value

        If Not dictChamadosPorCidade(cidade).exists(tipoServico) Then
            dictChamadosPorCidade(cidade)(tipoServico) = 0
        End If

        dictChamadosPorCidade(cidade)(tipoServico) = dictChamadosPorCidade(cidade)(tipoServico) + 1
    Next i

    ' Definir a planilha de resultados
    On Error Resume Next
    Set resultadosWs = ThisWorkbook.Sheets("Resultados")
    If resultadosWs Is Nothing Then
        Set resultadosWs = ThisWorkbook.Sheets.Add
        resultadosWs.Name = "Resultados"
    End If
    On Error GoTo 0

    ' Limpar conteúdo da planilha Resultados
    resultadosWs.Cells.Clear
    ' Definir o período dos dados
    periodo = Format(minDate, "dd/mm/yyyy") & " a " & Format(maxDate, "dd/mm/yyyy")
    ' Organizar os resultados em tabelas dinâmicas
    CriarTabela resultadosWs, "TIPO DE ATENDIMENTO", dictChamadosNovosPorServico, 1, False, periodo
    CriarTabela resultadosWs, "FINALIZADOS POR SERVIÇO", dictChamadosFinalizados, resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2, False, periodo

    ' Tabelas principais
    CriarTabela resultadosWs, "NOVAS + D0", dictStatus, resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2, False, periodo

    ' Tabela de SLA
    Dim dictSLA As Object
    Set dictSLA = CreateObject("Scripting.Dictionary")
    dictSLA.Add "Dentro do SLA", dentroSLA
    dictSLA.Add "Fora do SLA", foraSLA
    CriarTabela resultadosWs, "SLA", dictSLA, resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2, False, periodo
' Criar a tabela com os dados de modelos
    CriarTabela resultadosWs, "Chamados por Modelo de Equipamento", dictModelos, resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2, False, periodo
    ' Tabela de Chamados Vencendo Hoje
    CriarTabela resultadosWs, "Chamados Vencendo Hoje", dictChamadosVencendoHoje, resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2, False, periodo
' Tabela de Chamados Finalizados no Mês
    CriarTabela resultadosWs, "Chamados Finalizados no Mês", dictChamadosFinalizadosMes, resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2, False, periodo

    ' Tabela de Chamados por Técnico Detalhados
    Dim tecnicoKey As Variant
    For Each tecnicoKey In dictTecnicosDetalhados.Keys
        CriarTabela resultadosWs, "Chamados por Técnico - " & tecnicoKey, dictTecnicosDetalhados(tecnicoKey), resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2, False, periodo
    Next tecnicoKey

    ' Criar tabelas para cada cidade
    Dim linhaInicial As Long
    linhaInicial = resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2

    Dim cidadeKey As Variant
    For Each cidadeKey In dictChamadosPorCidade.Keys
        CriarTabela resultadosWs, "Chamados por Tipo de Serviço - " & cidadeKey, dictChamadosPorCidade(cidadeKey), linhaInicial, False, periodo
        linhaInicial = resultadosWs.Cells(resultadosWs.Rows.Count, 1).End(xlUp).Row + 2
    Next cidadeKey

    ' Excluir a planilha temporária
    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True

    MsgBox "Painel de informações detalhado criado com sucesso!"
End Sub

Sub CriarTabela(ws As Worksheet, nomeTabela As String, dados As Object, linhaInicial As Long, Optional csvImport As Boolean = False, Optional periodo As String = "")
    Dim i As Long
    Dim lastRow As Long
    Dim tempWs As Worksheet

    ' Verificar e criar planilhas se não existirem
    On Error Resume Next
    Set resultadosWs = ThisWorkbook.Sheets("Resultados")
    If resultadosWs Is Nothing Then
        Set resultadosWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets.Count)
        resultadosWs.Name = "Resultados"
    End If

    Set graficosWs = ThisWorkbook.Sheets("Graficos")
    If graficosWs Is Nothing Then
        Set graficosWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets.Count)
        graficosWs.Name = "Graficos"
    End If
    On Error GoTo 0

    If csvImport Then
        ' Criar a planilha temporária
        On Error Resume Next
        Set tempWs = Worksheets("Temp")
        On Error GoTo 0
        If tempWs Is Nothing Then
            Set tempWs = Worksheets.Add
            tempWs.Name = "Temp"
        End If

        ' Importar o arquivo CSV para a planilha temporária
        Dim csvFile As String
        Dim qt As QueryTable
        csvFile = Application.GetOpenFilename("Arquivos CSV (*.csv), *.csv", , "Selecione o arquivo CSV")
        If csvFile = "False" Then Exit Sub

        Set qt = tempWs.QueryTables.Add(Connection:="TEXT;" & csvFile, Destination:=tempWs.Range("A1"))
        With qt
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFilePlatform = xlWindows
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .Refresh BackgroundQuery:=False
        End With

        ' Converter colunas de data para o formato datetime
        tempWs.Columns("P:Q").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        tempWs.Columns("S:S").NumberFormat = "dd/mm/yyyy hh:mm:ss"

        ' Encontrar a última linha com dados
        lastRow = tempWs.Cells(tempWs.Rows.Count, "A").End(xlUp).Row
    End If

    ' Adicionar título com período
    ws.Cells(linhaInicial, 1).Value = nomeTabela & " (" & periodo & ")"
    ws.Cells(linhaInicial, 1).Font.Underline = xlUnderlineStyleSingle
    ws.Cells(linhaInicial, 1).Font.Bold = True
    ws.Cells(linhaInicial, 1).Font.Size = 12
    ws.Cells(linhaInicial, 1).HorizontalAlignment = xlCenter

    ' Definir cabeçalhos
    ws.Cells(linhaInicial + 1, 1).Value = "Serviço / Número do Chamado"
    ws.Cells(linhaInicial + 1, 2).Value = "Quantidade / Motivo"
    ws.Rows(linhaInicial + 1).Font.Bold = True
    ws.Rows(linhaInicial + 1).HorizontalAlignment = xlCenter

    ' Preencher os dados
    i = linhaInicial + 2

    Dim Key As Variant
    For Each Key In dados.Keys
        ws.Cells(i, 1).Value = Key
        If IsArray(dados(Key)) Then
            ws.Cells(i, 2).Value = dados(Key)(1) ' Mostrar motivo de cancelamento
        Else
            ws.Cells(i, 2).Value = dados(Key)
        End If
        i = i + 1
    Next Key

    ' Ajustar a largura das colunas
    ws.Columns("A:B").AutoFit

    ' Verificar se há dados para criar a tabela
    If i > linhaInicial + 2 Then
        ' Criar a tabela formatada
        With ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(linhaInicial + 1, 1), ws.Cells(i - 1, 2)), , xlYes)
            .Name = nomeTabela
            .TableStyle = "TableStyleMedium9"
            .ShowTotals = True
        End With
    Else
        MsgBox "Nenhum dado encontrado para criar a tabela " & nomeTabela, vbExclamation
    End If

    ' Excluir a planilha temporária, se criada
    If csvImport Then
        Application.DisplayAlerts = False
        tempWs.Delete
        Application.DisplayAlerts = True
    End If
End Sub

