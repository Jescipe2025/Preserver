#Include "TOTVS.ch"
#Include "TopConn.ch"


//+-------------------------------------------------------------------------+
//| Programa  | XFLUXJOB | Autor | Josue Izidio        | Data | 29/10/2026  |
//+-------------------------------------------------------------------------+
//| Descricao | JOB - Projecao do orcamento (SE7) pelo faturamento.          |
//|           | Tipo de calculo pelo parametro MV_XFLXTIP (M / C / U).        |
//+-------------------------------------------------------------------------+


/*/{Protheus.doc} XFLUXJOB
Job (sem interface) de projecao de fluxo de caixa por cliente na SE7.
Processa a EMPRESA/FILIAL em que o job foi executado.
A selecao de empresas/filiais (todas as empresas, todas as filiais etc.)
e feita no proprio cadastro do Schedule do Protheus.

  a) Le o faturamento mensal dos ultimos 12 meses fechados (mes atual excluido)
  b) Calcula a projecao conforme o parametro MV_XFLXTIP:
       M = Media     : soma dos 12 meses fechados / 12, igual em todos os meses
       C = Correspondente: cada mes recebe o faturamento do MESMO mes do
                       ultimo ano fechado (ex.: jan/2025 -> jan/2026).
                       Para os meses do ano seguinte cujo mes de 12 meses atras
                       ainda nao fechou, usa o mesmo mes de 24 meses atras.
       U = Ultimo mes: realizado do ultimo mes fechado em todos os meses
                       (em outubro usa setembro; em novembro usa outubro)
     Parametro vazio ou invalido: assume M (comportamento original).
  c) Grava a projecao nos meses do ano corrente (a partir do mes atual) e em
     todos os meses do ano seguinte (E7_VALJAN1..E7_VALDEZ1)
  d) Considera apenas clientes com A1_XFLUCX = 'S'
  e) Grava o realizado de cada mes nos campos E7_XREAJAN..E7_XREADEZ
  f) Pesquisa SE7 por FILIAL + CLIENTE + ANO: se encontrar ALTERA, senao INCLUI

Seguranca: se a empresa/filial nao tiver os campos customizados criados
(SA1 e SE7), o job apenas registra no console e encerra, sem erro.

Execucao:
  - Schedule do Protheus (SIGACFG > Schedule) -> rotina U_XFLUXJOB,
    selecionando no agendamento as empresas/filiais desejadas
  - Ou appserver.ini [ONSTART] com nParms=2 / Parm1=empresa / Parm2=filial
  - Ou StartJob("U_XFLUXJOB", GetEnvServer(), .F., {"04","01"})

@type User Function
@param xParam1, array|character, {cEmpresa, cFilial} (Schedule/StartJob) ou cEmpresa (ONSTART)
@param xParam2, character, cFilial quando chamado via ONSTART com nParms
/*/

// Nomes de campos customizados (Protheus aceita no maximo 10 caracteres)
#Define CAMPO_CLI   "E7_XCLIENT"   // Codigo-Loja do cliente (ex.: 000151-01)
#Define CAMPO_CGC   "E7_XCGC"      // CPF/CNPJ do cliente  (campo "Cgc. Cliente") - CONFIRMAR NOME NO SX3
#Define CAMPO_NOME  "E7_XNOMCLI"   // Nome do cliente
#Define CAMPO_DNAT  "E7_DESCNAT"   // Descricao da natureza (campo "Desc Naturez") - CONFIRMAR NOME NO SX3
#Define CAMPO_MOED  "E7_MOEDA"     // Moeda
#Define MOEDA_PAD   "1"            // Moeda gravada nos registros (campo caractere)
#Define NOME_JOB    "XFLUXJOB"
#Define VERSAO_JOB  "2026-10-30 v4"  // aparece no console.log para conferir o RPO
#Define PARAM_TIPO  "MV_XFLXTIP"     // Tipo de projecao: M (media), C (mes correspondente), U (ultimo mes)

User Function XFLUXJOB(xParam1, xParam2)
    Local lPrepEnv := .F.
    Local cEmp     := ""
    Local cFil     := ""

    // Se o ambiente nao vier aberto (ex.: ONSTART/StartJob), abre com os parametros
    If Type("cEmpAnt") <> "C"
        If ValType(xParam1) == "A" .And. Len(xParam1) >= 2     // {cEmp, cFil}
            cEmp := xParam1[1]
            cFil := xParam1[2]
        ElseIf ValType(xParam1) == "C" .And. ValType(xParam2) == "C" // ONSTART Parm1/Parm2
            cEmp := xParam1
            cFil := xParam2
        EndIf

        If Empty(cEmp) .Or. Empty(cFil)
            ConOut("[" + NOME_JOB + "] Ambiente nao preparado e empresa/filial nao informadas.")
            Return
        EndIf

        RpcSetType(3) // nao consome licenca
        If !RpcSetEnv(cEmp, cFil, , , "FIN", NOME_JOB, {"SA1", "SF2", "SE7"})
            ConOut("[" + NOME_JOB + "] Falha ao abrir ambiente " + cEmp + "/" + cFil)
            Return
        EndIf
        lPrepEnv := .T.
    EndIf

    // Empresa/filial sem os campos customizados: nao processa (evita erro de SQL)
    If !CamposOK()
        If lPrepEnv
            RpcClearEnv()
        EndIf
        Return
    EndIf

    // Trava por empresa + filial: filiais diferentes podem rodar em paralelo,
    // mas a mesma filial nao roda duas vezes ao mesmo tempo
    If LockByName(NOME_JOB, .T., .T.)
        XLog("Inicio do processamento - versao " + VERSAO_JOB + ;
             " - anos gerados: " + StrZero(Year(Date()), 4) + " e " + StrZero(Year(Date()) + 1, 4))

        If Empty(xFilial("SE7")) .And. !Empty(xFilial("SF2"))
            XLog("ATENCAO: SE7 compartilhada e SF2 exclusiva - valores de uma filial " + ;
                 "serao sobrescritos pela proxima. Verifique o compartilhamento.")
        EndIf

        XFlxProc()

        XLog("Fim do processamento")
        UnLockByName(NOME_JOB, .T., .T.)
    Else
        XLog("Rotina ja esta em execucao para esta empresa/filial. Execucao abortada.")
    EndIf

    If lPrepEnv
        RpcClearEnv()
    EndIf
Return

//-------------------------------------------------------------------
// Verifica se os campos customizados existem nesta empresa
//-------------------------------------------------------------------
Static Function CamposOK()
    Local aFalta := {}
    Local aCpos  := {}
    Local nI     := 0

    DbSelectArea("SA1")
    If SA1->(FieldPos("A1_XFLUCX")) == 0
        AAdd(aFalta, "A1_XFLUCX")
    EndIf

    DbSelectArea("SE7")
    aCpos := {CAMPO_CLI, CAMPO_CGC, CAMPO_NOME}
    AEval(CposRea(), {|x| AAdd(aCpos, x)})
    For nI := 1 To Len(aCpos)
        If SE7->(FieldPos(aCpos[nI])) == 0
            AAdd(aFalta, aCpos[nI])
        EndIf
    Next nI

    If Len(aFalta) > 0
        XLog("Empresa/filial ignorada - campos inexistentes: " + ArrTokStr(aFalta, ", "))
        Return .F.
    EndIf
Return .T.

//-------------------------------------------------------------------
// Processamento da empresa/filial corrente
//-------------------------------------------------------------------
Static Function XFlxProc()
    Local dRef    := Date()
    Local cIni    := DToS(MonthSub(FirstDay(dRef), 12)) // 1o dia, 12 meses atras
    Local cFim    := DToS(LastDay(dRef) + 1)            // 1o dia do PROXIMO mes (exclusivo):
                                                        // inclui o mes atual (parcial) para o realizado
    Local cMesAtu := Left(DToS(dRef), 6)                // AAAAMM do mes atual (fora da media)
    Local cMesAnt := Left(DToS(FirstDay(dRef) - 1), 6)  // AAAAMM do ultimo mes fechado (tipo U)
    Local cTipo   := TipoProj()                         // M, C ou U
    Local nAnoAtu := Year(dRef)
    Local nMesAtu := Month(dRef)
    Local cFilSE7 := xFilial("SE7")
    Local cQuery  := ""
    Local cAlias  := ""
    Local cCgc    := ""
    Local cNome   := ""
    Local aMeses  := {}
    Local nTotal  := 0
    Local nMedia  := 0
    Local nCli    := 0

    XLog("Tipo de projecao (" + PARAM_TIPO + "): " + cTipo + " - " + DescTipo(cTipo, cMesAnt))
    XLog("Periodo base: " + cIni + " a " + DToS(dRef - Day(dRef)) + ;
         " | Realizado ate: " + DToS(dRef) + " (mes atual parcial)")

    // a) Faturamento mensal por cliente (apenas A1_XFLUCX = 'S')
    //    ATENCAO: ajuste o campo de valor e os filtros da SF2 para ficarem
    //    iguais aos da sua query original (ex.: F2_VALFAT, CFOP, duplicata etc.)
    cQuery := " SELECT SA1.A1_CGC CGC, "
    cQuery += "        MAX(SA1.A1_NOME) NOME, "
    cQuery += "        SUBSTRING(SF2.F2_EMISSAO, 1, 6) ANOMES, "
    cQuery += "        SUM(SF2.F2_VALBRUT) VALOR "
    cQuery += " FROM " + RetSqlName("SF2") + " SF2 "
    cQuery += " INNER JOIN " + RetSqlName("SA1") + " SA1 "
    cQuery += "    ON SA1.A1_FILIAL  = '" + xFilial("SA1") + "' "
    cQuery += "   AND SA1.A1_COD     = SF2.F2_CLIENTE "
    cQuery += "   AND SA1.A1_LOJA    = SF2.F2_LOJA "
    cQuery += "   AND SA1.A1_XFLUCX  = 'S' "
    cQuery += "   AND SA1.A1_CGC    <> ' ' "
    cQuery += "   AND SA1.D_E_L_E_T_ = ' ' "
    cQuery += " WHERE SF2.F2_FILIAL  = '" + xFilial("SF2") + "' "
    cQuery += "   AND SF2.F2_TIPO    = 'N' "
    cQuery += "   AND SF2.F2_EMISSAO >= '" + cIni + "' "
    cQuery += "   AND SF2.F2_EMISSAO <  '" + cFim + "' "
    cQuery += "   AND SF2.D_E_L_E_T_ = ' ' "
    cQuery += " GROUP BY SA1.A1_CGC, SUBSTRING(SF2.F2_EMISSAO, 1, 6) "
    cQuery += " ORDER BY SA1.A1_CGC, SUBSTRING(SF2.F2_EMISSAO, 1, 6) "

    cAlias := MPSysOpenQuery(cQuery)

    While !(cAlias)->(Eof())
        cCgc   := (cAlias)->CGC
        cNome  := (cAlias)->NOME
        aMeses := {}
        nTotal := 0

        While !(cAlias)->(Eof()) .And. (cAlias)->CGC == cCgc
            AAdd(aMeses, {(cAlias)->ANOMES, (cAlias)->VALOR})
            // A media usa apenas os 12 meses fechados (mes atual parcial fica fora)
            If (cAlias)->ANOMES < cMesAtu
                nTotal += (cAlias)->VALOR
            EndIf
            (cAlias)->(DbSkip())
        EndDo

        // b) Soma dos 12 meses / 12 (usada no tipo M e sempre exibida no log)
        nMedia := Round(nTotal / 12, 2)

        // c) e e) Grava projecao (conforme o tipo) e realizado
        AtuCliente(cFilSE7, cCgc, AllTrim(cNome), aMeses, nMedia, cIni, nAnoAtu, nMesAtu, ;
                   cTipo, cMesAtu, cMesAnt)

        XLog("Cliente " + AllTrim(cCgc) + ;
             " - Total 12m: " + AllTrim(Transform(nTotal, "@E 999,999,999.99")) + ;
             " - Media: " + AllTrim(Transform(nMedia, "@E 999,999,999.99")) + ;
             IIf(cTipo == "U", " - Ultimo mes (" + cMesAnt + "): " + ;
                 AllTrim(Transform(ValMes(aMeses, cMesAnt), "@E 999,999,999.99")), ""))
        nCli++
    EndDo

    (cAlias)->(DbCloseArea())

    XLog("Clientes processados: " + cValToChar(nCli))
Return

//-------------------------------------------------------------------
// Monta os campos por ano e grava a SE7 do cliente
//-------------------------------------------------------------------
Static Function AtuCliente(cFilSE7, cCgc, cNome, aMeses, nMedia, cIni, nAnoAtu, nMesAtu, ;
                           cTipo, cMesAtu, cMesAnt)
    Local aCpoOrc := CposOrc()
    Local aCpoRea := CposRea()
    Local aAnos   := {} // { {cAno, { {cCampo, nValor}, ... } }, ... }
    Local cAnoMes := ""
    Local nPos    := 0
    Local nValor  := 0
    Local nMes    := 0
    Local nI      := 0

    // Realizado: somente o ANO CORRENTE, de janeiro ate o mes atual
    // (o mes atual entra parcial). Mes sem venda grava zero.
    // Nao gera registro do ano anterior.
    For nMes := 1 To nMesAtu
        cAnoMes := StrZero(nAnoAtu, 4) + StrZero(nMes, 2)
        nPos    := AScan(aMeses, {|x| x[1] == cAnoMes})
        nValor  := IIf(nPos > 0, aMeses[nPos][2], 0)
        AddCampo(aAnos, StrZero(nAnoAtu, 4), aCpoRea[nMes], nValor)
    Next nMes

    // Projecao ano corrente: do mes atual ate dezembro (meses passados nao sao alterados)
    For nMes := nMesAtu To 12
        AddCampo(aAnos, StrZero(nAnoAtu, 4), aCpoOrc[nMes], ;
                 ValorProj(cTipo, nMes, aMeses, nMedia, cMesAtu, cMesAnt))
    Next nMes

    // Projecao ano seguinte: janeiro a dezembro
    For nMes := 1 To 12
        AddCampo(aAnos, StrZero(nAnoAtu + 1, 4), aCpoOrc[nMes], ;
                 ValorProj(cTipo, nMes, aMeses, nMedia, cMesAtu, cMesAnt))
    Next nMes

    Begin Transaction
        For nI := 1 To Len(aAnos)
            GravaSE7(cFilSE7, cCgc, cNome, aAnos[nI][1], aAnos[nI][2])
        Next nI
    End Transaction
Return

//-------------------------------------------------------------------
// Pesquisa FILIAL + CLIENTE + ANO: se encontrar ALTERA, senao INCLUI
//-------------------------------------------------------------------
Static Function GravaSE7(cFilSE7, cCgc, cNome, cAno, aCampos)
    Local nRec    := BuscaSE7(cFilSE7, cCgc, cAno)
    Local lNovo   := (nRec == 0)
    Local cNatur  := SuperGetMV("MV_XNATFLX", .F., "")
    Local cDesNat := ""
    Local cCodLoj := CodLojaCli(cCgc)
    Local nI      := 0

    // Natureza: remove aspas/espacos e ajusta ao tamanho do campo
    cNatur := AllTrim(StrTran(StrTran(cNatur, '"', ""), "'", ""))
    If !Empty(cNatur)
        cNatur  := PadR(cNatur, TamSX3("E7_NATUREZ")[1])
        cDesNat := AllTrim(Posicione("SED", 1, xFilial("SED") + cNatur, "ED_DESCRIC"))
        If Empty(cDesNat)
            XLog("Natureza [" + AllTrim(cNatur) + "] nao encontrada na SED (verifique MV_XNATFLX).")
        EndIf
    EndIf

    DbSelectArea("SE7")
    If !lNovo
        SE7->(DbGoTo(nRec))           // localizou -> alteracao
    EndIf

    If RecLock("SE7", lNovo)          // lNovo = .T. inclui / .F. altera
        If lNovo
            SetCampo("E7_FILIAL", cFilSE7)
            SetCampo("E7_ANO"   , cAno)
        EndIf

        // Gravados na inclusao e na alteracao (corrige registros antigos)
        SetCampo(CAMPO_CLI , cCodLoj)          // Codigo-Loja do SA1
        SetCampo(CAMPO_CGC , cCgc)             // CPF/CNPJ
        SetCampo(CAMPO_MOED, MOEDA_PAD)        // Moeda 1
        If !Empty(cNatur)
            SetCampo("E7_NATUREZ", cNatur)
            // Descricao: grava so se o campo for real (virtual e preenchido pelo inicializador)
            If SE7->(FieldPos(CAMPO_DNAT)) > 0
                SetCampo(CAMPO_DNAT, cDesNat)
            EndIf
        EndIf

        SetCampo(CAMPO_NOME, cNome)

        For nI := 1 To Len(aCampos)
            SetCampo(aCampos[nI][1], aCampos[nI][2])
        Next nI

        SE7->(MsUnlock())

        XLog("SE7 " + IIf(lNovo, "INCLUIDO", "ALTERADO") + ;
             " - Cliente " + cCodLoj + " (" + AllTrim(cCgc) + ") Ano " + cAno)
    Else
        XLog("Nao foi possivel travar SE7 - Cliente " + AllTrim(cCgc) + " Ano " + cAno)
    EndIf
Return

//-------------------------------------------------------------------
// Localiza o RECNO da SE7 por FILIAL + CPF/CNPJ + ANO (0 = nao existe)
// Tambem encontra registros antigos, gravados com o CNPJ em E7_XCLIENT,
// para que sejam convertidos (alterados) em vez de duplicados.
//-------------------------------------------------------------------
Static Function BuscaSE7(cFilSE7, cCgc, cAno)
    Local cQry := ""
    Local cAli := ""
    Local nRec := 0

    cQry := " SELECT R_E_C_N_O_ REC "
    cQry += " FROM " + RetSqlName("SE7")
    cQry += " WHERE E7_FILIAL  = '" + cFilSE7 + "' "
    cQry += "   AND ( " + CAMPO_CGC + " = '" + cCgc + "' "
    cQry += "      OR " + CAMPO_CLI + " = '" + cCgc + "' ) "
    cQry += "   AND E7_ANO     = '" + cAno + "' "
    cQry += "   AND D_E_L_E_T_ = ' ' "
    cQry += " ORDER BY R_E_C_N_O_ "

    cAli := MPSysOpenQuery(cQry)
    If !(cAli)->(Eof())
        nRec := (cAli)->REC
    EndIf
    (cAli)->(DbCloseArea())
Return nRec

//-------------------------------------------------------------------
// Grava um campo da SE7 validando se existe no dicionario
//-------------------------------------------------------------------
Static Function SetCampo(cCampo, xValor)
    Local nPos := SE7->(FieldPos(cCampo))

    If nPos > 0
        // Ajusta o tipo ao do campo (ex.: moeda numerica ou caractere)
        If ValType(SE7->(FieldGet(nPos))) == "C" .And. ValType(xValor) == "N"
            xValor := cValToChar(xValor)
        ElseIf ValType(SE7->(FieldGet(nPos))) == "N" .And. ValType(xValor) == "C"
            xValor := Val(xValor)
        EndIf
        SE7->(FieldPut(nPos, xValor))
    Else
        XLog("Campo nao encontrado na SE7: " + cCampo)
    EndIf
Return

//-------------------------------------------------------------------
// Retorna "CODIGO-LOJA" do cliente no SA1 a partir do CPF/CNPJ
// (indice 3 do SA1: A1_FILIAL + A1_CGC). Se houver mais de uma loja
// com o mesmo CNPJ, prioriza a que tem A1_XFLUCX = 'S'.
//-------------------------------------------------------------------
Static Function CodLojaCli(cCgc)
    Local aArea  := SA1->(GetArea())
    Local cFilA1 := xFilial("SA1")
    Local cRet   := ""

    SA1->(DbSetOrder(3))
    If SA1->(DbSeek(cFilA1 + cCgc))
        cRet := AllTrim(SA1->A1_COD) + "-" + AllTrim(SA1->A1_LOJA)
        While !SA1->(Eof()) .And. SA1->A1_FILIAL == cFilA1 .And. AllTrim(SA1->A1_CGC) == AllTrim(cCgc)
            If SA1->A1_XFLUCX == "S"
                cRet := AllTrim(SA1->A1_COD) + "-" + AllTrim(SA1->A1_LOJA)
                Exit
            EndIf
            SA1->(DbSkip())
        EndDo
    Else
        XLog("Cliente com CNPJ " + AllTrim(cCgc) + " nao localizado no SA1 (indice 3).")
    EndIf

    RestArea(aArea)
Return cRet

//-------------------------------------------------------------------
// Le o parametro MV_XFLXTIP. Vazio ou invalido -> "M" (comportamento original)
//-------------------------------------------------------------------
Static Function TipoProj()
    Local cTipo := Upper(AllTrim(SuperGetMV(PARAM_TIPO, .F., "M")))

    cTipo := StrTran(StrTran(cTipo, '"', ""), "'", "")
    If Len(cTipo) != 1 .Or. !(cTipo $ "MCU")
        XLog("Parametro " + PARAM_TIPO + " com conteudo invalido [" + cTipo + "]. Usando M (media).")
        cTipo := "M"
    EndIf
Return cTipo

//-------------------------------------------------------------------
// Texto do tipo para o log
//-------------------------------------------------------------------
Static Function DescTipo(cTipo, cMesAnt)
    Local cDesc := ""

    Do Case
        Case cTipo == "C"
            cDesc := "mes correspondente do ultimo ano fechado"
        Case cTipo == "U"
            cDesc := "realizado do ultimo mes fechado (" + Right(cMesAnt, 2) + "/" + Left(cMesAnt, 4) + ") em todos os meses"
        OtherWise
            cDesc := "media dos ultimos 12 meses fechados"
    EndCase
Return cDesc

//-------------------------------------------------------------------
// Valor projetado para um mes (1..12), conforme o tipo
//   aMeses : { {AAAAMM, nValor}, ... } faturamento do cliente
//   M -> media dos 12 meses fechados
//   C -> faturamento do mesmo mes dentro dos 12 meses fechados.
//        A janela de 12 meses fechados tem cada mes do ano uma unica vez,
//        entao sempre e o mes mais recente ja fechado:
//          out/2026 -> out/2025 | jan/2027 -> jan/2026 | out/2027 -> out/2025
//   U -> realizado do ultimo mes fechado
//   Mes sem faturamento retorna zero.
//-------------------------------------------------------------------
Static Function ValorProj(cTipo, nMes, aMeses, nMedia, cMesAtu, cMesAnt)
    Local nValor := 0
    Local nPos   := 0
    Local cMM    := StrZero(nMes, 2)

    Do Case
        Case cTipo == "C"
            nPos := AScan(aMeses, {|x| Right(x[1], 2) == cMM .And. x[1] < cMesAtu})
            nValor := IIf(nPos > 0, aMeses[nPos][2], 0)

        Case cTipo == "U"
            nValor := ValMes(aMeses, cMesAnt)

        OtherWise   // "M"
            nValor := nMedia
    EndCase
Return Round(nValor, 2)

//-------------------------------------------------------------------
// Faturamento de um AAAAMM especifico (zero se nao houver)
//-------------------------------------------------------------------
Static Function ValMes(aMeses, cAnoMes)
    Local nPos := AScan(aMeses, {|x| x[1] == cAnoMes})
Return IIf(nPos > 0, aMeses[nPos][2], 0)

//-------------------------------------------------------------------
// Acumula campo/valor no ano correspondente
//-------------------------------------------------------------------
Static Function AddCampo(aAnos, cAno, cCampo, nValor)
    Local nPos := AScan(aAnos, {|x| x[1] == cAno})

    If nPos == 0
        AAdd(aAnos, {cAno, {}})
        nPos := Len(aAnos)
    EndIf
    AAdd(aAnos[nPos][2], {cCampo, nValor})
Return

// Campos de orcamento (projecao)
Static Function CposOrc()
Return {"E7_VALJAN1", "E7_VALFEV1", "E7_VALMAR1", "E7_VALABR1", "E7_VALMAI1", "E7_VALJUN1", ;
        "E7_VALJUL1", "E7_VALAGO1", "E7_VALSET1", "E7_VALOUT1", "E7_VALNOV1", "E7_VALDEZ1"}

// Campos de realizado (customizados)
Static Function CposRea()
Return {"E7_XREAJAN", "E7_XREAFEV", "E7_XREAMAR", "E7_XREAABR", "E7_XREAMAI", "E7_XREAJUN", ;
        "E7_XREAJUL", "E7_XREAAGO", "E7_XREASET", "E7_XREAOUT", "E7_XREANOV", "E7_XREADEZ"}

// Log no console do AppServer (identifica empresa/filial)
Static Function XLog(cMsg)
    ConOut("[" + NOME_JOB + "][" + cEmpAnt + "/" + AllTrim(cFilAnt) + "] " + ;
           DToC(Date()) + " " + Time() + " - " + cMsg)
Return

// Permite cadastrar a rotina no Schedule do Protheus
Static Function SchedDef()
Return {"P", "PARAMDEF", "", {}, ""}
