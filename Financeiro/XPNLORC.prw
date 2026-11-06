#Include "TOTVS.ch"

/*/{Protheus.doc} XPNLORC
Gera tres planilhas Excel (base para o painel de Orcamentos x Clientes),
reunindo todas as empresas do grupo em uma unica aba por planilha:
  1) Clientes.xls       - SA1, somente clientes com orcamento na SE7
                          nos anos informados (E7_XCLIENT = A1_COD + '-' + A1_LOJA)
  2) Orcamentos.xls     - SE7 somente dos anos informados (Ano 1 e Ano 2),
                          uma linha por natureza/mes (Orcado x Realizado)
  3) NotasFiscais.xls   - SF2 emitidas somente nos anos informados e
                          somente de clientes com orcamento na SE7
Empresas lidas: 01 a 10 (ver fPoEmpresas). Empresa sem a tabela e ignorada.
Chamar pelo menu: U_XPNLORC()
/*/
User Function XPNLORC()

    Local oDlg      := NIL
    Local oGetPasta := NIL
    Local cPasta    := PadR( GetTempPath(), 200 )
    Local cAno1     := StrZero( Year( Date() ) - 1, 4 )
    Local cAno2     := StrZero( Year( Date() ), 4 )
    Local lOk       := .F.
    Local aEmp      := fPoEmpresas()
    Local nCli      := 0
    Local nOrc      := 0
    Local nNfs      := 0
    Local cLink     := AllTrim( SuperGetMV( "MV_LINKORC", .F., "" ) )

    DEFINE MSDIALOG oDlg TITLE "Orcamentos x Clientes - Gerar Excel" ;
        FROM 0, 0 TO 215, 480 PIXEL

        @ 12, 20  SAY "Ano 1:" SIZE 40, 10 PIXEL
        @ 22, 20  MSGET cAno1 SIZE 30, 12 PICTURE "9999" PIXEL

        @ 12, 70  SAY "Ano 2:" SIZE 40, 10 PIXEL
        @ 22, 70  MSGET cAno2 SIZE 30, 12 PICTURE "9999" PIXEL

        @ 42, 20  SAY "Pasta de destino:" SIZE 170, 10 PIXEL
        @ 52, 20  MSGET oGetPasta VAR cPasta SIZE 195, 12 PIXEL

        @ 52, 220 BUTTON "..." SIZE 18, 12 PIXEL ;
            ACTION ( cPasta := PadR( fPoEscolhePasta( AllTrim( cPasta ) ), 200 ), ;
                     oGetPasta:Refresh() )

        @ 80, 10  BUTTON "&Gerar"    SIZE 55, 15 PIXEL ;
            ACTION ( iif( fPoValida( cAno1, cAno2, cPasta ), ( lOk := .T., oDlg:End() ), NIL ) )
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

    MsgRun( "Gerando planilha de clientes, aguarde...", "Processando", ;
            { || nCli := fPoGerar( cPasta + "Clientes.xls", "Clientes", "TB01", ;
                                   fPoColClientes(), fPoQryClientes( aEmp, cAno1, cAno2 ) ) } )

    MsgRun( "Gerando planilha de orcamentos, aguarde...", "Processando", ;
            { || nOrc := fPoGerar( cPasta + "Orcamentos.xls", "Orcamentos", "TB02", ;
                                   fPoColOrcamentos(), fPoQryOrcamentos( aEmp, cAno1, cAno2 ) ) } )

    MsgRun( "Gerando planilha de notas fiscais, aguarde...", "Processando", ;
            { || nNfs := fPoGerar( cPasta + "NotasFiscais.xls", "Notas Fiscais", "TB03", ;
                                   fPoColNotas(), fPoQryNotas( aEmp, cAno1, cAno2 ) ) } )

    MsgInfo( "Planilhas geradas em " + cPasta + CRLF + CRLF + ;
             "Clientes: "       + cValToChar( nCli ) + " linhas" + CRLF + ;
             "Orcamentos: "     + cValToChar( nOrc ) + " linhas" + CRLF + ;
             "Notas fiscais: "  + cValToChar( nNfs ) + " linhas", "Concluido" )

    If MsgYesNo( "Deseja visualizar o Painel de Orcamentos x Clientes?", "Painel" )
        If Empty( cLink )
            MsgAlert( "Parametro MV_LINKORC nao configurado. Informe o link do Painel.", "Atencao" )
        Else
            WinExec( 'cmd.exe /c start "" "' + cLink + '"' )
        EndIf
    EndIf

Return NIL

/*/{Protheus.doc} fPoEmpresas
Empresas do grupo: { codigo, nome }. A tabela de cada uma e <ALIAS><COD>0,
por exemplo SA1010, SE7020, SF2100.
@type   Static Function
@return Array
/*/
Static Function fPoEmpresas()
//--------------------------------------------------------------------
Return { { "01", "PRESERVE TRANSP"      }, ;
         { "02", "LISERVE VIGILAN"      }, ;
         { "03", "LISERVE SERVICO"      }, ;
         { "04", "LISERVE SERV E TERC"  }, ;
         { "05", "PRESERVE/PB TRA"      }, ;
         { "06", "DELFO SEGURANCA"      }, ;
         { "07", "OPCAO SEGURANCA"      }, ;
         { "08", "TAXI AEREO PRES"      }, ;
         { "09", "PARTINVEST"           }, ;
         { "10", "TDPG EMPR IMOBILIARIOS E HOTELEIROS LTDA" } }

/*/{Protheus.doc} fPoValida
Valida os parametros da tela.
@type   Static Function
@return Logical
/*/
Static Function fPoValida( cAno1, cAno2, cPasta )
//--------------------------------------------------------------------
    If Len( AllTrim( cAno1 ) ) != 4 .Or. Val( cAno1 ) < 1900
        MsgAlert( "Informe o Ano 1 com 4 digitos.", "Atencao" )
        Return .F.
    EndIf
    If Len( AllTrim( cAno2 ) ) != 4 .Or. Val( cAno2 ) < 1900
        MsgAlert( "Informe o Ano 2 com 4 digitos.", "Atencao" )
        Return .F.
    EndIf
    If Empty( cPasta )
        MsgAlert( "Informe a pasta de destino.", "Atencao" )
        Return .F.
    EndIf
Return .T.

//====================================================================
// QUERIES  (uma UNION ALL por empresa que possui a tabela)
//====================================================================

/*/{Protheus.doc} fPoQryClientes
SA1 de todas as empresas, somente clientes que possuem orcamento na SE7
da mesma empresa nos anos informados (ligacao: E7_XCLIENT = A1_COD + '-' + A1_LOJA).
/*/
Static Function fPoQryClientes( aEmp, cAno1, cAno2 )
//--------------------------------------------------------------------
    Local cQuery := ""
    Local cTab   := ""
    Local cSe7   := ""
    Local nI     := 0

    For nI := 1 To Len( aEmp )
        cTab := "SA1" + aEmp[ nI, 1 ] + "0"
        cSe7 := "SE7" + aEmp[ nI, 1 ] + "0"
        If !TCCanOpen( cTab ) .Or. !TCCanOpen( cSe7 )
            Loop
        EndIf
        If !Empty( cQuery )
            cQuery += " UNION ALL "
        EndIf
        cQuery += "SELECT '" + aEmp[ nI, 1 ] + "' AS COD_EMP, " + ;
                  "'" + aEmp[ nI, 2 ] + "' AS EMPRESA, " + ;
                  "A1_FILIAL, A1_COD, A1_LOJA, A1_PESSOA, A1_CGC, " + ;
                  "A1_NOME, A1_NREDUZ, A1_EST, A1_MUN, A1_ULTCOM, A1_RISCO " + ;
                  "FROM " + cTab + " A1 " + ;
                  "WHERE A1.D_E_L_E_T_ <> '*' " + ;
                  "AND EXISTS ( SELECT 1 FROM " + cSe7 + " E7 " + ;
                  "WHERE E7.E7_XCLIENT = A1.A1_COD + '-' + A1.A1_LOJA " + ;
                  "AND E7.D_E_L_E_T_ <> '*' "

        // Filtro de ano: comente a linha abaixo para trazer clientes
        // com orcamento em qualquer ano
        cQuery += "AND E7.E7_ANO IN ('" + cAno1 + "','" + cAno2 + "') "

        cQuery += ") "
    Next nI

    If !Empty( cQuery )
        cQuery += " ORDER BY COD_EMP, A1_COD, A1_LOJA"
    EndIf

Return cQuery

/*/{Protheus.doc} fPoQryOrcamentos
SE7 dos dois anos informados, uma linha por natureza e mes.
A descricao da natureza vem da SED da mesma empresa.
/*/
Static Function fPoQryOrcamentos( aEmp, cAno1, cAno2 )
//--------------------------------------------------------------------
    Local cQuery := ""
    Local cTab   := ""
    Local cSed   := ""
    Local cDesc  := ""
    Local cApply := ""
    Local nI     := 0

    For nI := 1 To Len( aEmp )
        cTab := "SE7" + aEmp[ nI, 1 ] + "0"
        cSed := "SED" + aEmp[ nI, 1 ] + "0"
        If !TCCanOpen( cTab )
            Loop
        EndIf

        If TCCanOpen( cSed )
            cDesc  := "ISNULL(ED.ED_DESCRIC, ' ') AS DESC_NAT, "
            cApply := "OUTER APPLY ( SELECT TOP 1 ED_DESCRIC FROM " + cSed + ;
                      " WHERE ED_CODIGO = E7.E7_NATUREZ AND D_E_L_E_T_ <> '*' ) ED "
        Else
            cDesc  := "' ' AS DESC_NAT, "
            cApply := ""
        EndIf

        If !Empty( cQuery )
            cQuery += " UNION ALL "
        EndIf
        cQuery += "SELECT '" + aEmp[ nI, 1 ] + "' AS COD_EMP, " + ;
                  "'" + aEmp[ nI, 2 ] + "' AS EMPRESA, " + ;
                  "E7.E7_FILIAL, E7.E7_ANO, M.MES, M.NOMEMES, " + ;
                  "E7.E7_NATUREZ, " + cDesc + ;
                  "E7.E7_XCGC, E7.E7_XCLIENT, E7.E7_XNOMCLI, " + ;
                  "M.ORCADO, M.REALIZADO, M.REALIZADO - M.ORCADO AS DIFERENCA " + ;
                  "FROM " + cTab + " E7 " + ;
                  "CROSS APPLY ( VALUES " + ;
                  "( 1,'JAN',E7.E7_VALJAN1,E7.E7_XREAJAN),( 2,'FEV',E7.E7_VALFEV1,E7.E7_XREAFEV)," + ;
                  "( 3,'MAR',E7.E7_VALMAR1,E7.E7_XREAMAR),( 4,'ABR',E7.E7_VALABR1,E7.E7_XREAABR)," + ;
                  "( 5,'MAI',E7.E7_VALMAI1,E7.E7_XREAMAI),( 6,'JUN',E7.E7_VALJUN1,E7.E7_XREAJUN)," + ;
                  "( 7,'JUL',E7.E7_VALJUL1,E7.E7_XREAJUL),( 8,'AGO',E7.E7_VALAGO1,E7.E7_XREAAGO)," + ;
                  "( 9,'SET',E7.E7_VALSET1,E7.E7_XREASET),(10,'OUT',E7.E7_VALOUT1,E7.E7_XREAOUT)," + ;
                  "(11,'NOV',E7.E7_VALNOV1,E7.E7_XREANOV),(12,'DEZ',E7.E7_VALDEZ1,E7.E7_XREADEZ) " + ;
                  ") M ( MES, NOMEMES, ORCADO, REALIZADO ) " + ;
                  cApply + ;
                  "WHERE E7.D_E_L_E_T_ <> '*' " + ;
                  "AND E7.E7_ANO IN ('" + cAno1 + "','" + cAno2 + "') "
    Next nI

    If !Empty( cQuery )
        cQuery += " ORDER BY COD_EMP, E7_FILIAL, E7_ANO, E7_NATUREZ, E7_XCLIENT, MES"
    EndIf

Return cQuery

/*/{Protheus.doc} fPoQryNotas
SF2 emitidas nos dois anos informados, nao deletadas, somente de clientes
com orcamento na SE7 da mesma empresa (E7_XCLIENT = F2_CLIENTE + '-' + F2_LOJA).
Tipos D (devolucao) e B
(beneficiamento) ficam de fora, pois o destinatario e fornecedor (SA2).
CNPJ e nome do cliente vem da SA1 da mesma empresa.
/*/
Static Function fPoQryNotas( aEmp, cAno1, cAno2 )
//--------------------------------------------------------------------
    Local cQuery := ""
    Local cTab   := ""
    Local cSa1   := ""
    Local cSe7   := ""
    Local cCli   := ""
    Local cApply := ""
    Local nI     := 0

    For nI := 1 To Len( aEmp )
        cTab := "SF2" + aEmp[ nI, 1 ] + "0"
        cSa1 := "SA1" + aEmp[ nI, 1 ] + "0"
        cSe7 := "SE7" + aEmp[ nI, 1 ] + "0"
        If !TCCanOpen( cTab ) .Or. !TCCanOpen( cSe7 )
            Loop
        EndIf

        If TCCanOpen( cSa1 )
            cCli   := "ISNULL(A1.A1_CGC, ' ') AS CGC, ISNULL(A1.A1_NOME, ' ') AS NOME, "
            cApply := "OUTER APPLY ( SELECT TOP 1 A1_CGC, A1_NOME FROM " + cSa1 + ;
                      " WHERE A1_COD = F2.F2_CLIENTE AND A1_LOJA = F2.F2_LOJA" + ;
                      " AND D_E_L_E_T_ <> '*' ) A1 "
        Else
            cCli   := "' ' AS CGC, ' ' AS NOME, "
            cApply := ""
        EndIf

        If !Empty( cQuery )
            cQuery += " UNION ALL "
        EndIf
        cQuery += "SELECT '" + aEmp[ nI, 1 ] + "' AS COD_EMP, " + ;
                  "'" + aEmp[ nI, 2 ] + "' AS EMPRESA, " + ;
                  "F2.F2_FILIAL, F2.F2_DOC, F2.F2_SERIE, F2.F2_EMISSAO, " + ;
                  "LEFT(F2.F2_EMISSAO, 4) AS ANO, SUBSTRING(F2.F2_EMISSAO, 5, 2) AS MES, " + ;
                  "F2.F2_TIPO, F2.F2_CLIENTE, F2.F2_LOJA, " + cCli + ;
                  "F2.F2_VALBRUT " + ;
                  "FROM " + cTab + " F2 " + ;
                  cApply + ;
                  "WHERE F2.D_E_L_E_T_ <> '*' " + ;
                  "AND F2.F2_TIPO NOT IN ('D','B') " + ;
                  "AND ( F2.F2_EMISSAO BETWEEN '" + cAno1 + "0101' AND '" + cAno1 + "1231' " + ;
                  "   OR F2.F2_EMISSAO BETWEEN '" + cAno2 + "0101' AND '" + cAno2 + "1231' ) " + ;
                  "AND EXISTS ( SELECT 1 FROM " + cSe7 + " E7 " + ;
                  "WHERE E7.E7_XCLIENT = F2.F2_CLIENTE + '-' + F2.F2_LOJA " + ;
                  "AND E7.D_E_L_E_T_ <> '*' "

        // Filtro de ano do orcamento: comente a linha abaixo para trazer
        // notas de clientes com orcamento em qualquer ano
        cQuery += "AND E7.E7_ANO IN ('" + cAno1 + "','" + cAno2 + "') "

        cQuery += ") "
    Next nI

    If !Empty( cQuery )
        cQuery += " ORDER BY COD_EMP, F2_FILIAL, F2_EMISSAO, F2_DOC, F2_SERIE"
    EndIf

Return cQuery

//====================================================================
// COLUNAS  { Titulo, Alinhamento(1=esq,2=centro,3=dir), Formato(1=geral,
//            2=numero,3=moeda,4=data), Tipo("C","N","D") }
// A ordem deve ser a mesma dos campos do SELECT.
//====================================================================

Static Function fPoColClientes()
//--------------------------------------------------------------------
Return { { "Cod. Empresa"   , 1, 1, "C" }, ;
         { "Empresa"        , 1, 1, "C" }, ;
         { "Filial"         , 1, 1, "C" }, ;
         { "Codigo"         , 1, 1, "C" }, ;
         { "Loja"           , 1, 1, "C" }, ;
         { "Pessoa"         , 1, 1, "C" }, ;
         { "CNPJ/CPF"       , 1, 1, "C" }, ;
         { "Razao Social"   , 1, 1, "C" }, ;
         { "Nome Fantasia"  , 1, 1, "C" }, ;
         { "UF"             , 1, 1, "C" }, ;
         { "Municipio"      , 1, 1, "C" }, ;
         { "Ult. Compra"    , 1, 1, "D" }, ;
         { "Risco"          , 1, 1, "C" } }

Static Function fPoColOrcamentos()
//--------------------------------------------------------------------
Return { { "Cod. Empresa"   , 1, 1, "C" }, ;
         { "Empresa"        , 1, 1, "C" }, ;
         { "Filial"         , 1, 1, "C" }, ;
         { "Ano"            , 1, 1, "C" }, ;
         { "Mes"            , 2, 1, "N" }, ;
         { "Nome Mes"       , 1, 1, "C" }, ;
         { "Natureza"       , 1, 1, "C" }, ;
         { "Desc. Natureza" , 1, 1, "C" }, ;
         { "CNPJ Cliente"   , 1, 1, "C" }, ;
         { "Cliente"        , 1, 1, "C" }, ;
         { "Nome Cliente"   , 1, 1, "C" }, ;
         { "Orcado"         , 3, 3, "N" }, ;
         { "Realizado"      , 3, 3, "N" }, ;
         { "Diferenca"      , 3, 3, "N" } }

Static Function fPoColNotas()
//--------------------------------------------------------------------
Return { { "Cod. Empresa"   , 1, 1, "C" }, ;
         { "Empresa"        , 1, 1, "C" }, ;
         { "Filial"         , 1, 1, "C" }, ;
         { "Nota"           , 1, 1, "C" }, ;
         { "Serie"          , 1, 1, "C" }, ;
         { "Emissao"        , 1, 1, "D" }, ;
         { "Ano"            , 1, 1, "C" }, ;
         { "Mes"            , 1, 1, "C" }, ;
         { "Tipo"           , 1, 1, "C" }, ;
         { "Cliente"        , 1, 1, "C" }, ;
         { "Loja"           , 1, 1, "C" }, ;
         { "CNPJ/CPF"       , 1, 1, "C" }, ;
         { "Razao Social"   , 1, 1, "C" }, ;
         { "Valor Bruto"    , 3, 3, "N" } }

//====================================================================
// GERACAO DO EXCEL
//====================================================================

/*/{Protheus.doc} fPoGerar
Executa a query e grava o resultado em uma aba do Excel.
@type   Static Function
@param  cArquivo, Char , Caminho completo do arquivo .xls
@param  cAba    , Char , Nome da aba
@param  cTab    , Char , Nome da tabela dentro da aba
@param  aCols   , Array, Definicao das colunas (ver bloco COLUNAS)
@param  cQuery  , Char , Query a executar (vazia = so cabecalho)
@return Numeric , Quantidade de linhas gravadas
/*/
Static Function fPoGerar( cArquivo, cAba, cTab, aCols, cQuery )
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
                AAdd( aRow, fPoFmt( (cAlia)->( FieldGet( nCol ) ), aCols[ nCol, 4 ] ) )
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

/*/{Protheus.doc} fPoFmt
Formata o valor conforme o tipo da coluna:
"N" numero; "D" data AAAAMMDD -> DD/MM/AAAA; "C" texto sem espacos.
/*/
Static Function fPoFmt( xVal, cTipo )
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

/*/{Protheus.doc} fPoEscolhePasta
Abre o dialogo de selecao de pasta. Se cancelado, retorna o caminho original.
/*/
Static Function fPoEscolhePasta( cAtual )
//--------------------------------------------------------------------
    Local cNova := cGetFile( "", "Selecione a pasta de destino", 1, cAtual, .T., ;
                             GETF_LOCALHARD + GETF_RETDIRECTORY )
Return iif( Empty( cNova ), cAtual, cNova )
