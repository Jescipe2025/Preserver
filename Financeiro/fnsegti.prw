#Include "TOTVS.ch"

//ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
//ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
//ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
//ฑฑบPrograma  ณ FNSEGTI บ Autor ณ Josu้ Izidio       บ Data ณ 05/11/2026 บฑฑ
//ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
//ฑฑบDescricao ณ Controle de Notas da TI                                     บฑฑ
//ฑฑฬออออออออออุอออออออออออออออัออออออออออออออออออออออออออออออออออออออออออออนฑฑ
//฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿


/*/{Protheus.doc} FNSEGTI
Gera duas planilhas Excel (base para o Painel Notas TI),
reunindo as empresas do grupo em uma unica aba por planilha:
  1) ContasPagar.xls  - SE2 com competencia (E2_YCOMPET) dentro do periodo,
                        lancados por ZARSON ALMEIDA ou aprovados por LEONARDO BARROS
  2) ItensNotas.xls   - SF1/SD1 com emissao dentro do periodo,
                        incluidas pelo usuario ZARSON ALMEIDA
Ambas trazem UF/municipio (A2_EST/A2_MUN) e o segmento do fornecedor (A2_XSEGMNT):
  1=Telecom; 2=Locacao; 3=Licenca Software; 4=Suporte TI; 5=Produtos
O contas a pagar traz o saldo (E2_SALDO) para o painel separar titulos
pagos, em aberto e pagos parcialmente.
Periodo: o usuario informa quantos meses deseja retroagir, contando o mes
atual. Ex.: em outubro/2026, 10 meses = 01/01/2026 a 31/10/2026.
Empresas lidas: ver fSgEmpresas. Empresa sem a tabela e ignorada.
Chamar pelo menu: U_FNSEGTI()
/*/
User Function FNSEGTI()

    Local oDlg      := NIL
    Local oGetPasta := NIL
    Local oSayPer   := NIL
    Local cPasta    := PadR( GetTempPath(), 200 )
    Local nMeses    := Month( Date() )
    Local cPeriodo  := ""
    Local lOk       := .F.
    Local aEmp      := fSgEmpresas()
    Local aPer      := {}
    Local nCpg      := 0
    Local nNfs      := 0
    Local cLink     := AllTrim( SuperGetMV( "MV_LINKSEG", .F., "" ) )

    cPeriodo := fSgTxtPer( nMeses )

    DEFINE MSDIALOG oDlg TITLE "Segmentos TI - Gerar Excel" ;
        FROM 0, 0 TO 215, 480 PIXEL

        @ 12, 20  SAY "Quantidade de meses a retroagir:" SIZE 120, 10 PIXEL
        @ 22, 20  MSGET nMeses SIZE 30, 12 PICTURE "99" PIXEL ;
            VALID ( cPeriodo := fSgTxtPer( nMeses ), oSayPer:Refresh(), .T. )

        @ 24, 60  SAY oSayPer VAR cPeriodo SIZE 170, 10 PIXEL

        @ 42, 20  SAY "Pasta de destino:" SIZE 170, 10 PIXEL
        @ 52, 20  MSGET oGetPasta VAR cPasta SIZE 195, 12 PIXEL

        @ 52, 220 BUTTON "..." SIZE 18, 12 PIXEL ;
            ACTION ( cPasta := PadR( fSgEscolhePasta( AllTrim( cPasta ) ), 200 ), ;
                     oGetPasta:Refresh() )

        @ 80, 10  BUTTON "&Gerar"    SIZE 55, 15 PIXEL ;
            ACTION ( iif( fSgValida( nMeses, cPasta ), ( lOk := .T., oDlg:End() ), NIL ) )
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

    // aPer: { DataIni AAAAMMDD, DataFim AAAAMMDD, AnoMesIni AAAAMM, AnoMesFim AAAAMM }
    aPer := fSgPeriodo( nMeses )

    MsgRun( "Gerando planilha de contas a pagar, aguarde...", "Processando", ;
            { || nCpg := fSgGerar( cPasta + "ContasPagar.xls", "Contas a Pagar", "TB01", ;
                                   fSgColCtasPagar(), fSgQryCtasPagar( aEmp, aPer ) ) } )

    MsgRun( "Gerando planilha de itens das notas, aguarde...", "Processando", ;
            { || nNfs := fSgGerar( cPasta + "ItensNotas.xls", "Itens Notas", "TB02", ;
                                   fSgColItensNotas(), fSgQryItensNotas( aEmp, aPer ) ) } )

    MsgInfo( "Planilhas geradas em " + cPasta + CRLF + ;
             "Periodo: " + fSgTxtPer( nMeses ) + CRLF + CRLF + ;
             "Contas a pagar: " + cValToChar( nCpg ) + " linhas" + CRLF + ;
             "Itens das notas: " + cValToChar( nNfs ) + " linhas", "Concluido" )

    If MsgYesNo( "Deseja visualizar o Painel Notas TI?", "Painel" )
        If Empty( cLink )
            MsgAlert( "Parametro MV_LINKSEG nao configurado. Informe o link do Painel.", "Atencao" )
        Else
            WinExec( 'cmd.exe /c start "" "' + cLink + '"' )
        EndIf
    EndIf

Return NIL

/*/{Protheus.doc} fSgEmpresas
Codigos das empresas do grupo. A tabela de cada uma e <ALIAS><COD>0,
por exemplo SE2010, SA2020, SF1050. O nome vem da SYS_COMPANY (M0_NOME).
CTT e SED sao compartilhadas (sempre CTT010 / SED010).
@type   Static Function
@return Array
/*/
Static Function fSgEmpresas()
//--------------------------------------------------------------------
Return { "01", "02", "05" }

/*/{Protheus.doc} fSgPeriodo
Calcula o periodo a partir da quantidade de meses (contando o mes atual).
@type   Static Function
@return Array { DataIni, DataFim, AnoMesIni, AnoMesFim }
/*/
Static Function fSgPeriodo( nMeses )
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

/*/{Protheus.doc} fSgTxtPer
Texto do periodo para exibicao na tela.
/*/
Static Function fSgTxtPer( nMeses )
//--------------------------------------------------------------------
    Local aPer := {}

    If nMeses < 1 .Or. nMeses > 60
        Return "Periodo invalido"
    EndIf

    aPer := fSgPeriodo( nMeses )

Return "De " + fSgFmt( aPer[ 1 ], "D" ) + " ate " + fSgFmt( aPer[ 2 ], "D" )

/*/{Protheus.doc} fSgValida
Valida os parametros da tela.
@type   Static Function
@return Logical
/*/
Static Function fSgValida( nMeses, cPasta )
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
// FILTROS FIXOS (altere aqui se mudarem os responsaveis)
//====================================================================

Static Function fSgFiltroCtas()
//--------------------------------------------------------------------
Return "AND ( UPPER(E2.E2_YUSER) = 'ZARSON ALMEIDA' OR UPPER(E2.E2_XAPROVA) = 'LEONARDO BARROS' ) "

Static Function fSgFiltroNotas()
//--------------------------------------------------------------------
Return "AND UPPER(F1.F1_USUARIO) = 'ZARSON ALMEIDA' "

/*/{Protheus.doc} fSgSegmento
Campos do segmento do fornecedor (codigo + descricao).
/*/
Static Function fSgSegmento()
//--------------------------------------------------------------------
Return "ISNULL(A2.A2_XSEGMNT, ' ') AS COD_SEGMENTO, " + ;
       "CASE A2.A2_XSEGMNT " + ;
       "WHEN '1' THEN 'TELECOM' " + ;
       "WHEN '2' THEN 'LOCACAO' " + ;
       "WHEN '3' THEN 'LICENCA SOFTWARE' " + ;
       "WHEN '4' THEN 'SUPORTE TI' " + ;
       "WHEN '5' THEN 'PRODUTOS' " + ;
       "ELSE 'NAO INFORMADO' END AS SEGMENTO, "

//====================================================================
// QUERIES  (uma UNION ALL por empresa que possui a tabela)
//====================================================================

/*/{Protheus.doc} fSgQryCtasPagar
SE2 de todas as empresas, competencia (E2_YCOMPET, formato MMAAAA)
dentro do periodo informado.
/*/
Static Function fSgQryCtasPagar( aEmp, aPer )
//--------------------------------------------------------------------
    Local cQuery := ""
    Local cSe2   := ""
    Local cSa2   := ""
    Local cComp  := "SUBSTRING(E2.E2_YCOMPET,3,4) + SUBSTRING(E2.E2_YCOMPET,1,2)"
    Local nI     := 0

    For nI := 1 To Len( aEmp )
        cSe2 := "SE2" + aEmp[ nI ] + "0"
        cSa2 := "SA2" + aEmp[ nI ] + "0"
        If !TCCanOpen( cSe2 ) .Or. !TCCanOpen( cSa2 )
            Loop
        EndIf
        If !Empty( cQuery )
            cQuery += " UNION ALL "
        EndIf
        cQuery += "SELECT '" + aEmp[ nI ] + "' AS COD_EMP, " + ;
                  "ISNULL(M0.M0_NOME, ' ') AS EMPRESA, ISNULL(M0.M0_FILIAL, ' ') AS NOME_FILIAL, " + ;
                  "E2.E2_PREFIXO, E2.E2_NUM, E2.E2_PARCELA, E2.E2_FORNECE, E2.E2_LOJA, " + ;
                  "LEFT(ISNULL(A2.A2_CGC, ' '), 8) AS RADIC_CNPJ, ISNULL(A2.A2_CGC, ' ') AS A2_CGC, " + ;
                  "ISNULL(A2.A2_NOME, ' ') AS A2_NOME, " + ;
                  "ISNULL(A2.A2_EST, ' ') AS A2_EST, ISNULL(A2.A2_MUN, ' ') AS A2_MUN, " + ;
                  fSgSegmento() + ;
                  "ISNULL(ED.ED_DESCRIC, ' ') AS ED_DESCRIC, " + ;
                  "E2.E2_EMISSAO, E2.E2_VENCTO, E2.E2_VALOR, E2.E2_SALDO, E2.E2_DTLANC, E2.E2_HIST, " + ;
                  "E2.E2_XAPROVA, E2.E2_YCOMPET, " + cComp + " AS ANOMES_COMP, E2.E2_BAIXA, " + ;
                  "E2.E2_CCD, ISNULL(CTT.CTT_DESC01, ' ') AS CTT_DESC01, E2.E2_YUSER, E2.E2_ORIGEM " + ;
                  "FROM " + cSe2 + " E2 " + ;
                  "LEFT OUTER JOIN SYS_COMPANY M0 ON M0.M0_CODIGO = '" + aEmp[ nI ] + "' " + ;
                  "AND E2.E2_FILIAL = M0.M0_CODFIL AND M0.D_E_L_E_T_ = ' ' " + ;
                  "LEFT OUTER JOIN " + cSa2 + " A2 ON E2.E2_FORNECE = A2.A2_COD " + ;
                  "AND E2.E2_LOJA = A2.A2_LOJA AND A2.D_E_L_E_T_ = ' ' " + ;
                  "LEFT OUTER JOIN CTT010 CTT ON E2.E2_CCD = CTT.CTT_CUSTO AND CTT.D_E_L_E_T_ = ' ' " + ;
                  "LEFT OUTER JOIN SED010 ED ON E2.E2_NATUREZ = ED.ED_CODIGO AND ED.D_E_L_E_T_ = ' ' " + ;
                  "WHERE E2.D_E_L_E_T_ = ' ' " + ;
                  "AND " + cComp + " BETWEEN '" + aPer[ 3 ] + "' AND '" + aPer[ 4 ] + "' " + ;
                  fSgFiltroCtas()
    Next nI

    If !Empty( cQuery )
        cQuery += " ORDER BY COD_EMP, NOME_FILIAL, ANOMES_COMP, E2_FORNECE, E2_NUM, E2_PARCELA"
    EndIf

Return cQuery

/*/{Protheus.doc} fSgQryItensNotas
SF1 + SD1 de todas as empresas, emissao dentro do periodo informado.
/*/
Static Function fSgQryItensNotas( aEmp, aPer )
//--------------------------------------------------------------------
    Local cQuery := ""
    Local cSf1   := ""
    Local cSd1   := ""
    Local cSa2   := ""
    Local cSb1   := ""
    Local nI     := 0

    For nI := 1 To Len( aEmp )
        cSf1 := "SF1" + aEmp[ nI ] + "0"
        cSd1 := "SD1" + aEmp[ nI ] + "0"
        cSa2 := "SA2" + aEmp[ nI ] + "0"
        cSb1 := "SB1" + aEmp[ nI ] + "0"
        If !TCCanOpen( cSf1 ) .Or. !TCCanOpen( cSd1 ) .Or. ;
           !TCCanOpen( cSa2 ) .Or. !TCCanOpen( cSb1 )
            Loop
        EndIf
        If !Empty( cQuery )
            cQuery += " UNION ALL "
        EndIf
        cQuery += "SELECT '" + aEmp[ nI ] + "' AS COD_EMP, " + ;
                  "ISNULL(M0.M0_NOME, ' ') AS EMPRESA, " + ;
                  "F1.F1_FILIAL, F1.F1_SERIE, F1.F1_DOC, F1.F1_FORNECE, F1.F1_LOJA, " + ;
                  "LEFT(ISNULL(A2.A2_CGC, ' '), 8) AS RADIC_CNPJ, ISNULL(A2.A2_CGC, ' ') AS A2_CGC, " + ;
                  "ISNULL(A2.A2_NOME, ' ') AS A2_NOME, " + ;
                  "ISNULL(A2.A2_EST, ' ') AS A2_EST, ISNULL(A2.A2_MUN, ' ') AS A2_MUN, " + ;
                  fSgSegmento() + ;
                  "F1.F1_EMISSAO, LEFT(F1.F1_EMISSAO, 4) AS ANO, SUBSTRING(F1.F1_EMISSAO, 5, 2) AS MES, " + ;
                  "ISNULL(D1.D1_COD, ' ') AS D1_COD, ISNULL(B1.B1_DESC, ' ') AS B1_DESC, " + ;
                  "CASE WHEN B1.B1_YSEMPC = 'N' THEN 'PRODUTO' ELSE 'SERVICO' END AS TIPO, " + ;
                  "ISNULL(D1.D1_QUANT, 0) AS D1_QUANT, ISNULL(D1.D1_VUNIT, 0) AS D1_VUNIT, " + ;
                  "ISNULL(D1.D1_TOTAL, 0) AS D1_TOTAL " + ;
                  "FROM " + cSf1 + " F1 " + ;
                  "LEFT OUTER JOIN SYS_COMPANY M0 ON M0.M0_CODIGO = '" + aEmp[ nI ] + "' " + ;
                  "AND F1.F1_FILIAL = M0.M0_CODFIL AND M0.D_E_L_E_T_ = ' ' " + ;
                  "LEFT OUTER JOIN " + cSd1 + " D1 ON F1.F1_FILIAL = D1.D1_FILIAL " + ;
                  "AND F1.F1_SERIE = D1.D1_SERIE AND F1.F1_DOC = D1.D1_DOC " + ;
                  "AND F1.F1_FORNECE = D1.D1_FORNECE AND F1.F1_LOJA = D1.D1_LOJA " + ;
                  "AND D1.D_E_L_E_T_ = ' ' " + ;
                  "LEFT OUTER JOIN " + cSa2 + " A2 ON F1.F1_FORNECE = A2.A2_COD " + ;
                  "AND F1.F1_LOJA = A2.A2_LOJA AND A2.D_E_L_E_T_ = ' ' " + ;
                  "LEFT OUTER JOIN " + cSb1 + " B1 ON D1.D1_COD = B1.B1_COD AND B1.D_E_L_E_T_ = ' ' " + ;
                  "WHERE F1.D_E_L_E_T_ = ' ' " + ;
                  "AND F1.F1_EMISSAO BETWEEN '" + aPer[ 1 ] + "' AND '" + aPer[ 2 ] + "' " + ;
                  fSgFiltroNotas()
    Next nI

    If !Empty( cQuery )
        cQuery += " ORDER BY COD_EMP, F1_FILIAL, F1_EMISSAO, F1_DOC, F1_SERIE, D1_COD"
    EndIf

Return cQuery

//====================================================================
// COLUNAS  { Titulo, Alinhamento(1=esq,2=centro,3=dir), Formato(1=geral,
//            2=numero,3=moeda,4=data), Tipo("C","N","D") }
// A ordem deve ser a mesma dos campos do SELECT.
//====================================================================

Static Function fSgColCtasPagar()
//--------------------------------------------------------------------
Return { { "Cod. Empresa"     , 1, 1, "C" }, ;
         { "Empresa"          , 1, 1, "C" }, ;
         { "Filial"           , 1, 1, "C" }, ;
         { "Prefixo"          , 1, 1, "C" }, ;
         { "Titulo"           , 1, 1, "C" }, ;
         { "Parcela"          , 1, 1, "C" }, ;
         { "Fornecedor"       , 1, 1, "C" }, ;
         { "Loja"             , 1, 1, "C" }, ;
         { "Radical CNPJ"     , 1, 1, "C" }, ;
         { "CNPJ/CPF"         , 1, 1, "C" }, ;
         { "Razao Social"     , 1, 1, "C" }, ;
         { "UF"               , 1, 1, "C" }, ;
         { "Municipio"        , 1, 1, "C" }, ;
         { "Cod. Segmento"    , 2, 1, "C" }, ;
         { "Segmento"         , 1, 1, "C" }, ;
         { "Natureza"         , 1, 1, "C" }, ;
         { "Emissao"          , 1, 1, "D" }, ;
         { "Vencimento"       , 1, 1, "D" }, ;
         { "Valor"            , 3, 3, "N" }, ;
         { "Saldo"            , 3, 3, "N" }, ;
         { "Dt. Lancamento"   , 1, 1, "D" }, ;
         { "Historico"        , 1, 1, "C" }, ;
         { "Aprovador"        , 1, 1, "C" }, ;
         { "Competencia"      , 1, 1, "C" }, ;
         { "Ano/Mes Compet."  , 1, 1, "C" }, ;
         { "Baixa"            , 1, 1, "D" }, ;
         { "Centro de Custo"  , 1, 1, "C" }, ;
         { "Desc. C. Custo"   , 1, 1, "C" }, ;
         { "Usuario"          , 1, 1, "C" }, ;
         { "Origem"           , 1, 1, "C" } }

Static Function fSgColItensNotas()
//--------------------------------------------------------------------
Return { { "Cod. Empresa"     , 1, 1, "C" }, ;
         { "Empresa"          , 1, 1, "C" }, ;
         { "Filial"           , 1, 1, "C" }, ;
         { "Serie"            , 1, 1, "C" }, ;
         { "Nota"             , 1, 1, "C" }, ;
         { "Fornecedor"       , 1, 1, "C" }, ;
         { "Loja"             , 1, 1, "C" }, ;
         { "Radical CNPJ"     , 1, 1, "C" }, ;
         { "CNPJ/CPF"         , 1, 1, "C" }, ;
         { "Razao Social"     , 1, 1, "C" }, ;
         { "UF"               , 1, 1, "C" }, ;
         { "Municipio"        , 1, 1, "C" }, ;
         { "Cod. Segmento"    , 2, 1, "C" }, ;
         { "Segmento"         , 1, 1, "C" }, ;
         { "Emissao"          , 1, 1, "D" }, ;
         { "Ano"              , 1, 1, "C" }, ;
         { "Mes"              , 1, 1, "C" }, ;
         { "Produto"          , 1, 1, "C" }, ;
         { "Descricao"        , 1, 1, "C" }, ;
         { "Tipo"             , 1, 1, "C" }, ;
         { "Quantidade"       , 3, 2, "N" }, ;
         { "Vlr. Unitario"    , 3, 3, "N" }, ;
         { "Vlr. Total"       , 3, 3, "N" } }

//====================================================================
// GERACAO DO EXCEL
//====================================================================

/*/{Protheus.doc} fSgGerar
Executa a query e grava o resultado em uma aba do Excel.
@type   Static Function
@param  cArquivo, Char , Caminho completo do arquivo .xls
@param  cAba    , Char , Nome da aba
@param  cTab    , Char , Nome da tabela dentro da aba
@param  aCols   , Array, Definicao das colunas (ver bloco COLUNAS)
@param  cQuery  , Char , Query a executar (vazia = so cabecalho)
@return Numeric , Quantidade de linhas gravadas
/*/
Static Function fSgGerar( cArquivo, cAba, cTab, aCols, cQuery )
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
                AAdd( aRow, fSgFmt( (cAlia)->( FieldGet( nCol ) ), aCols[ nCol, 4 ] ) )
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

/*/{Protheus.doc} fSgFmt
Formata o valor conforme o tipo da coluna:
"N" numero; "D" data AAAAMMDD -> DD/MM/AAAA; "C" texto sem espacos.
/*/
Static Function fSgFmt( xVal, cTipo )
//--------------------------------------------------------------------
    Local cVal := ""

    If cTipo == "N"
        Return iif( ValType( xVal ) == "N", xVal, Val( cValToChar( xVal ) ) )
    EndIf

    If ValType( xVal ) == "D"
        Return DtoC( xVal )
    EndIf

    cVal := AllTrim( cValToChar( xVal ) )

    If cTipo == "D"
        If Len( cVal ) == 8
            Return SubStr( cVal, 7, 2 ) + "/" + SubStr( cVal, 5, 2 ) + "/" + SubStr( cVal, 1, 4 )
        EndIf
        Return ""
    EndIf

Return cVal

/*/{Protheus.doc} fSgEscolhePasta
Abre o dialogo de selecao de pasta. Se cancelado, retorna o caminho original.
/*/
Static Function fSgEscolhePasta( cAtual )
//--------------------------------------------------------------------
    Local cNova := cGetFile( "", "Selecione a pasta de destino", 1, cAtual, .T., ;
                             GETF_LOCALHARD + GETF_RETDIRECTORY )
Return iif( Empty( cNova ), cAtual, cNova )
