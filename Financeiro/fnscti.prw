#Include "TOTVS.ch"

//�����������������������������������������������������������������������������
//�������������������������������������������������������������������������ͻ��
//���Programa  � FNSCTI  � Autor � Josue Izidio       � Data � 06/11/2026 ���
//�������������������������������������������������������������������������͹��
//���Descricao � Rastreio das Solicitacoes de Compra da TI (SC x Workflow   ���
//���          � x Pedido)                                                  ���
//�������������������������������������������������������������������������ͼ��
//�����������������������������������������������������������������������������


/*/{Protheus.doc} FNSCTI
Gera uma planilha Excel (base para o Painel SC TI) reunindo as empresas
do grupo em uma unica aba:
  SolicitacoesCompra.xls - SC1 com emissao dentro do periodo, dos usuarios
                           definidos em fScFiltroUsr, com o status do
                           workflow (WFA) e os dados do pedido gerado (SC7/SA2).
Sao exportadas todas as SCs (com e sem pedido); a separacao entre elas
e feita no proprio painel.
Periodo: o usuario informa quantos meses deseja retroagir, contando o mes
atual. Ex.: em novembro/2026, 11 meses = 01/01/2026 a 30/11/2026.
Do workflow (WFA) e trazido apenas o registro mais recente de cada processo,
para nao duplicar itens da SC.
Empresas lidas: ver fScEmpresas. Empresa sem a SC1 e ignorada; sem WFA ou
SC7 a planilha sai com essas colunas em branco.
Chamar pelo menu: U_FNSCTI()
/*/
User Function FNSCTI()

    Local oDlg      := NIL
    Local oGetPasta := NIL
    Local oSayPer   := NIL
    Local cPasta    := PadR( GetTempPath(), 200 )
    Local nMeses    := Month( Date() )
    Local cPeriodo  := ""
    Local lOk       := .F.
    Local aEmp      := fScEmpresas()
    Local aPer      := {}
    Local nLin      := 0
    Local cArq      := ""
    Local cLink     := AllTrim( SuperGetMV( "MV_LINKSC", .F., "" ) )

    cPeriodo := fScTxtPer( nMeses )

    DEFINE MSDIALOG oDlg TITLE "Solicitacoes de Compra TI - Gerar Excel" ;
        FROM 0, 0 TO 215, 480 PIXEL

        @ 12, 20  SAY "Quantidade de meses a retroagir:" SIZE 120, 10 PIXEL
        @ 22, 20  MSGET nMeses SIZE 30, 12 PICTURE "99" PIXEL ;
            VALID ( cPeriodo := fScTxtPer( nMeses ), oSayPer:Refresh(), .T. )

        @ 24, 60  SAY oSayPer VAR cPeriodo SIZE 170, 10 PIXEL

        @ 42, 20  SAY "Pasta de destino:" SIZE 170, 10 PIXEL
        @ 52, 20  MSGET oGetPasta VAR cPasta SIZE 195, 12 PIXEL

        @ 52, 220 BUTTON "..." SIZE 18, 12 PIXEL ;
            ACTION ( cPasta := PadR( fScEscolhePasta( AllTrim( cPasta ) ), 200 ), ;
                     oGetPasta:Refresh() )

        @ 80, 10  BUTTON "&Gerar"    SIZE 55, 15 PIXEL ;
            ACTION ( iif( fScValida( nMeses, cPasta ), ( lOk := .T., oDlg:End() ), NIL ) )
        @ 80, 80  BUTTON "&Cancelar" SIZE 55, 15 PIXEL ;
            ACTION oDlg:End()

    ACTIVATE MSDIALOG oDlg CENTERED

    If !lOk
        Return NIL
    EndIf

    cPasta := AllTrim( cPasta )
    If Right( cPasta, 1 ) != "\"
        cPasta += "\"
    EndIf
    cArq := cPasta + "SolicitacoesCompra.xls"

    // aPer: { DataIni AAAAMMDD, DataFim AAAAMMDD, AnoMesIni AAAAMM, AnoMesFim AAAAMM }
    aPer := fScPeriodo( nMeses )

    MsgRun( "Gerando planilha de solicitacoes de compra, aguarde...", "Processando", ;
            { || nLin := fScGerar( cArq, "Solicitacoes", "TB01", ;
                                   fScColunas(), fScQuery( aEmp, aPer ) ) } )

    MsgInfo( "Planilha gerada em " + cArq + CRLF + ;
             "Periodo: " + fScTxtPer( nMeses ) + CRLF + ;
             CRLF + ;
             "Itens de SC: " + cValToChar( nLin ) + " linhas", "Concluido" )

    If MsgYesNo( "Deseja visualizar o Painel SC TI?", "Painel" )
        If Empty( cLink )
            MsgAlert( "Parametro MV_LINKSC nao configurado. Informe o link do Painel.", "Atencao" )
        Else
            WinExec( 'cmd.exe /c start "" "' + cLink + '"' )
        EndIf
    EndIf

Return NIL

/*/{Protheus.doc} fScEmpresas
Codigos das empresas do grupo. As tabelas de cada uma sao <ALIAS><COD>0,
por exemplo SC1010, WFA020, SC7050. O nome vem da SYS_COMPANY (M0_NOME).
@type   Static Function
@return Array
/*/
Static Function fScEmpresas()
//--------------------------------------------------------------------
Return { "01", "02", "04", "05", "09" }

/*/{Protheus.doc} fScPeriodo
Calcula o periodo a partir da quantidade de meses (contando o mes atual).
@type   Static Function
@return Array { DataIni, DataFim, AnoMesIni, AnoMesFim }
/*/
Static Function fScPeriodo( nMeses )
//--------------------------------------------------------------------
    Local nAno := Year( Date() )
    Local nMes := Month( Date() ) - ( nMeses - 1 )
    Local cIni := ""
    Local cFim := DtoS( LastDay( Date() ) )

    While nMes < 1
        nMes += 12
        nAno--
    EndDo

    cIni := StrZero( nAno, 4 ) + StrZero( nMes, 2 ) + "01"

Return { cIni, cFim, Left( cIni, 6 ), Left( cFim, 6 ) }

/*/{Protheus.doc} fScTxtPer
Texto do periodo para exibicao na tela.
/*/
Static Function fScTxtPer( nMeses )
//--------------------------------------------------------------------
    Local aPer := {}

    If nMeses < 1 .Or. nMeses > 60
        Return "Periodo invalido"
    EndIf

    aPer := fScPeriodo( nMeses )

Return "De " + fScFmt( aPer[ 1 ], "D" ) + " ate " + fScFmt( aPer[ 2 ], "D" )

/*/{Protheus.doc} fScValida
Valida os parametros da tela.
@type   Static Function
@return Logical
/*/
Static Function fScValida( nMeses, cPasta )
//--------------------------------------------------------------------
    If nMeses < 1 .Or. nMeses > 60
        MsgAlert( "Informe a quantidade de meses entre 1 e 60.", "Atencao" )
        Return .F.
    EndIf
    If Empty( cPasta )
        MsgAlert( "Informe a pasta de destino.", "Atencao" )
        Return .F.
    EndIf
Return .T.

//====================================================================
// FILTROS FIXOS (altere aqui se mudarem os solicitantes)
//====================================================================

Static Function fScFiltroUsr()
//--------------------------------------------------------------------
Return "AND C1.C1_USER IN ('001124', '000954') "

/*/{Protheus.doc} fScStatusWf
Descricao do status do workflow a partir do WFA_TIPO.
/*/
Static Function fScStatusWf( cTipo )
//--------------------------------------------------------------------
Return "CASE " + ;
       "WHEN " + cTipo + " = '0' THEN 'EMAIL SENDO PROCESSADO' " + ;
       "WHEN " + cTipo + " = '1' THEN 'EMAIL NA CAIXA DE ENTRADA (INBOX)' " + ;
       "WHEN " + cTipo + " = '2' THEN 'EMAIL NA CAIXA DE SAIDA (OUTBOX)' " + ;
       "WHEN " + cTipo + " = '3' THEN 'EMAIL ENVIADO (SENT)' " + ;
       "WHEN " + cTipo + " = '4' THEN 'EMAIL ARQUIVADO (ARCHIVE)' " + ;
       "WHEN " + cTipo + " = '5' THEN 'EMAIL ENVIADO COM ERRO (OUTBOX/ERROR)' " + ;
       "WHEN " + cTipo + " = '6' THEN 'EMAIL RECEBIDO COM ERRO (INBOX/ERROR)' " + ;
       "WHEN " + cTipo + " = '7' THEN 'EMAIL IGNORADO (IGNORED)' " + ;
       "WHEN " + cTipo + " = '8' THEN 'EMAIL DE WORKFLOW VIA-LINK ENVIADO' " + ;
       "WHEN C1.C1_WFID = ' ' THEN 'SEM WORKFLOW' " + ;
       "ELSE 'NAO LOCALIZADO NO WORKFLOW' END"

//====================================================================
// QUERY  (uma UNION ALL por empresa que possui a SC1)
//====================================================================

/*/{Protheus.doc} fScQuery
SC1 de todas as empresas, emissao dentro do periodo informado, com o
ultimo status do workflow (WFA) e o item do pedido gerado (SC7 + SA2).
/*/
Static Function fScQuery( aEmp, aPer )
//--------------------------------------------------------------------
    Local cQuery   := ""
    Local cEmp     := ""
    Local cSc1     := ""
    Local cWfa     := ""
    Local cSc7     := ""
    Local cSa2     := ""
    Local cTipo    := ""
    Local cCpoWf   := ""
    Local cJoinWf  := ""
    Local cCpoPed  := ""
    Local cJoinPed := ""
    Local lPed     := .F.
    Local nI       := 0

    For nI := 1 To Len( aEmp )
        cEmp := aEmp[ nI ]
        cSc1 := "SC1" + cEmp + "0"
        cWfa := "WFA" + cEmp + "0"
        cSc7 := "SC7" + cEmp + "0"
        cSa2 := "SA2" + cEmp + "0"

        If !TCCanOpen( cSc1 )
            Loop
        EndIf

        // Workflow: so o registro mais recente do processo (evita duplicar itens)
        If TCCanOpen( cWfa )
            cTipo   := "ISNULL(WF.WFA_TIPO, ' ')"
            cJoinWf := "OUTER APPLY ( SELECT TOP 1 W.WFA_TIPO, W.WFA_DATA, W.WFA_HORA " + ;
                       "FROM " + cWfa + " W " + ;
                       "WHERE C1.C1_WFPROC <> ' ' " + ;
                       "AND W.WFA_IDENT = REPLACE(C1.C1_WFPROC, '.htm', '') " + ;
                       "ORDER BY W.WFA_DATA DESC, W.WFA_HORA DESC ) WF "
            cCpoWf  := cTipo + " AS WFA_TIPO, " + fScStatusWf( cTipo ) + " AS STATUS_WF, " + ;
                       "ISNULL(WF.WFA_DATA, ' ') AS WFA_DATA, ISNULL(WF.WFA_HORA, ' ') AS WFA_HORA, "
        Else
            cJoinWf := ""
            cCpoWf  := "' ' AS WFA_TIPO, " + fScStatusWf( "' '" ) + " AS STATUS_WF, " + ;
                       "' ' AS WFA_DATA, ' ' AS WFA_HORA, "
        EndIf

        // Pedido de compra gerado pela SC
        lPed := TCCanOpen( cSc7 )
        If lPed
            cJoinPed := "LEFT OUTER JOIN " + cSc7 + " C7 ON C7.C7_FILIAL = C1.C1_FILIAL " + ;
                        "AND C7.C7_NUM = C1.C1_PEDIDO AND C7.C7_ITEM = C1.C1_ITEMPED " + ;
                        "AND C1.C1_PEDIDO <> ' ' AND C7.D_E_L_E_T_ = ' ' "
            cCpoPed  := "ISNULL(C7.C7_EMISSAO, ' ') AS C7_EMISSAO, ISNULL(C7.C7_FORNECE, ' ') AS C7_FORNECE, " + ;
                        "ISNULL(C7.C7_LOJA, ' ') AS C7_LOJA, "
            If TCCanOpen( cSa2 )
                cJoinPed += "LEFT OUTER JOIN " + cSa2 + " A2 ON A2.A2_COD = C7.C7_FORNECE " + ;
                            "AND A2.A2_LOJA = C7.C7_LOJA AND A2.D_E_L_E_T_ = ' ' "
                cCpoPed  += "ISNULL(A2.A2_NOME, ' ') AS A2_NOME, "
            Else
                cCpoPed  += "' ' AS A2_NOME, "
            EndIf
            cCpoPed += "ISNULL(C7.C7_QUANT, 0) AS C7_QUANT, ISNULL(C7.C7_PRECO, 0) AS C7_PRECO, " + ;
                       "ISNULL(C7.C7_TOTAL, 0) AS C7_TOTAL, ISNULL(C7.C7_QUJE, 0) AS C7_QUJE "
        Else
            cJoinPed := ""
            cCpoPed  := "' ' AS C7_EMISSAO, ' ' AS C7_FORNECE, ' ' AS C7_LOJA, ' ' AS A2_NOME, " + ;
                        "0 AS C7_QUANT, 0 AS C7_PRECO, 0 AS C7_TOTAL, 0 AS C7_QUJE "
        EndIf

        If !Empty( cQuery )
            cQuery += " UNION ALL "
        EndIf

        cQuery += "SELECT '" + cEmp + "' AS COD_EMP, " + ;
                  "ISNULL(M0.M0_NOME, ' ') AS EMPRESA, ISNULL(M0.M0_FILIAL, ' ') AS NOME_FILIAL, " + ;
                  "C1.C1_FILIAL, C1.C1_NUM, C1.C1_ITEM, C1.C1_PRODUTO, C1.C1_GRPRD, C1.C1_DESCRI, C1.C1_UM, " + ;
                  "C1.C1_QUANT, C1.C1_QUJE, C1.C1_SOLICIT, C1.C1_YAPROVA, C1.C1_APROV, C1.C1_YDTAPRO, " + ;
                  "C1.C1_NOMAPRO, C1.C1_OBS, C1.C1_EMISSAO, C1.C1_DATPRF, C1.C1_RESIDUO, C1.C1_COTACAO, " + ;
                  "C1.C1_PEDIDO, C1.C1_ITEMPED, " + ;
                  "CASE WHEN C1.C1_PEDIDO <> ' ' THEN 'COM PEDIDO' ELSE 'SEM PEDIDO' END AS SITUACAO_PED, " + ;
                  "C1.C1_WFID, C1.C1_WFPROC, C1.C1_WFDAT, C1.C1_WFHOR, " + ;
                  cCpoWf + cCpoPed + ;
                  "FROM " + cSc1 + " C1 " + ;
                  "LEFT OUTER JOIN SYS_COMPANY M0 ON M0.M0_CODIGO = '" + cEmp + "' " + ;
                  "AND C1.C1_FILIAL = M0.M0_CODFIL AND M0.D_E_L_E_T_ = ' ' " + ;
                  cJoinWf + cJoinPed + ;
                  "WHERE C1.D_E_L_E_T_ = ' ' " + ;
                  "AND C1.C1_EMISSAO BETWEEN '" + aPer[ 1 ] + "' AND '" + aPer[ 2 ] + "' " + ;
                  fScFiltroUsr()
    Next nI

    If !Empty( cQuery )
        cQuery += " ORDER BY COD_EMP, C1_FILIAL, C1_EMISSAO, C1_NUM, C1_ITEM"
    EndIf

Return cQuery

//====================================================================
// COLUNAS  { Titulo, Alinhamento(1=esq,2=centro,3=dir), Formato(1=geral,
//            2=numero,3=moeda,4=data), Tipo("C","N","D") }
// A ordem deve ser a mesma dos campos do SELECT.
// Os titulos sao lidos pelo Painel SC TI: nao altere sem ajustar o painel.
//====================================================================

Static Function fScColunas()
//--------------------------------------------------------------------
Return { { "Cod. Empresa"     , 1, 1, "C" }, ;
         { "Empresa"          , 1, 1, "C" }, ;
         { "Nome Filial"      , 1, 1, "C" }, ;
         { "Filial"           , 1, 1, "C" }, ;
         { "Num. SC"          , 1, 1, "C" }, ;
         { "Item"             , 1, 1, "C" }, ;
         { "Produto"          , 1, 1, "C" }, ;
         { "Grupo"            , 1, 1, "C" }, ;
         { "Descricao"        , 1, 1, "C" }, ;
         { "UM"               , 1, 1, "C" }, ;
         { "Quantidade"       , 3, 2, "N" }, ;
         { "Qtd. Atendida"    , 3, 2, "N" }, ;
         { "Solicitante"      , 1, 1, "C" }, ;
         { "Aprovador SC"     , 1, 1, "C" }, ;
         { "Status Aprov."    , 2, 1, "C" }, ;
         { "Dt. Aprovacao"    , 1, 1, "D" }, ;
         { "Nome Aprovador"   , 1, 1, "C" }, ;
         { "Observacao"       , 1, 1, "C" }, ;
         { "Emissao"          , 1, 1, "D" }, ;
         { "Necessidade"      , 1, 1, "D" }, ;
         { "Residuo"          , 2, 1, "C" }, ;
         { "Cotacao"          , 1, 1, "C" }, ;
         { "Pedido"           , 1, 1, "C" }, ;
         { "Item Pedido"      , 1, 1, "C" }, ;
         { "Situacao Pedido"  , 1, 1, "C" }, ;
         { "WF ID"            , 1, 1, "C" }, ;
         { "WF Processo"      , 1, 1, "C" }, ;
         { "WF Data"          , 1, 1, "D" }, ;
         { "WF Hora"          , 1, 1, "C" }, ;
         { "WF Tipo"          , 2, 1, "C" }, ;
         { "Status WF"        , 1, 1, "C" }, ;
         { "WFA Data"         , 1, 1, "D" }, ;
         { "WFA Hora"         , 1, 1, "C" }, ;
         { "Emissao Pedido"   , 1, 1, "D" }, ;
         { "Fornecedor"       , 1, 1, "C" }, ;
         { "Loja"             , 1, 1, "C" }, ;
         { "Razao Social"     , 1, 1, "C" }, ;
         { "Qtd. Pedido"      , 3, 2, "N" }, ;
         { "Preco Unit."      , 3, 3, "N" }, ;
         { "Total Pedido"     , 3, 3, "N" }, ;
         { "Qtd. Entregue"    , 3, 2, "N" } }

//====================================================================
// GERACAO DO EXCEL
//====================================================================

/*/{Protheus.doc} fScGerar
Executa a query e grava o resultado em uma aba do Excel.
@type   Static Function
@param  cArquivo, Char , Caminho completo do arquivo .xls
@param  cAba    , Char , Nome da aba
@param  cTab    , Char , Nome da tabela dentro da aba
@param  aCols   , Array, Definicao das colunas (ver bloco COLUNAS)
@param  cQuery  , Char , Query a executar (vazia = so cabecalho)
@return Numeric , Quantidade de linhas gravadas
/*/
Static Function fScGerar( cArquivo, cAba, cTab, aCols, cQuery )
//--------------------------------------------------------------------
    Local oFWMsExcel := FWMSExcel():New()
    Local cAlia      := GetNextAlias()
    Local nCols      := Len( aCols )
    Local nCol       := 0
    Local nReg       := 0
    Local aRow       := {}

    oFWMsExcel:AddWorkSheet( cAba )
    oFWMsExcel:AddTable( cAba, cTab )

    For nCol := 1 To nCols
        oFWMsExcel:AddColumn( cAba, cTab, aCols[ nCol, 1 ], aCols[ nCol, 2 ], aCols[ nCol, 3 ], .F. )
    Next nCol

    If !Empty( cQuery )
        DbUseArea( .T., "TOPCONN", TcGenQry( , , cQuery ), cAlia, .F., .T. )

        (cAlia)->( DbGoTop() )
        While !(cAlia)->( EOF() )
            aRow := {}
            For nCol := 1 To nCols
                AAdd( aRow, fScFmt( (cAlia)->( FieldGet( nCol ) ), aCols[ nCol, 4 ] ) )
            Next nCol
            oFWMsExcel:AddRow( cAba, cTab, aRow )
            nReg++
            (cAlia)->( DbSkip() )
        EndDo
        (cAlia)->( DbCloseArea() )
    EndIf

    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile( cArquivo )
    oFWMsExcel:DeActivate()

Return nReg

/*/{Protheus.doc} fScFmt
Formata o valor conforme o tipo da coluna:
"N" numero; "D" data AAAAMMDD -> DD/MM/AAAA; "C" texto sem espacos.
/*/
Static Function fScFmt( xVal, cTipo )
//--------------------------------------------------------------------
    Local cVal := ""

    If cTipo == "N"
        Return iif( ValType( xVal ) == "N", xVal, Val( cValToChar( xVal ) ) )
    EndIf

    If ValType( xVal ) == "D"
        Return iif( Empty( xVal ), "", DtoC( xVal ) )
    EndIf

    cVal := AllTrim( cValToChar( xVal ) )

    If cTipo == "D"
        If Len( cVal ) == 8
            Return SubStr( cVal, 7, 2 ) + "/" + SubStr( cVal, 5, 2 ) + "/" + SubStr( cVal, 1, 4 )
        EndIf
        Return ""
    EndIf

Return cVal

/*/{Protheus.doc} fScEscolhePasta
Abre o dialogo de selecao de pasta. Se cancelado, retorna o caminho original.
/*/
Static Function fScEscolhePasta( cAtual )
//--------------------------------------------------------------------
    Local cNova := cGetFile( "", "Selecione a pasta de destino", 1, cAtual, .T., ;
                             GETF_LOCALHARD + GETF_RETDIRECTORY )
Return iif( Empty( cNova ), cAtual, cNova )
