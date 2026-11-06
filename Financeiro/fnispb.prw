#INCLUDE "TOTVS.CH"

/*/{Protheus.doc} User Function FNispb
    Busca o ISPB de um banco cadastrado em um arquivo JSON na pasta system.
    @type  Function
    @since 28/10/2026
    @param cCodBanco, Character, Codigo do banco (Ex: "422")
    @return cIspb, Character, ISPB encontrado ou vazio
/*/
User Function FNispb(cCodBanco)

	Local cJsonPath := ""
	Local cJsonData := ""
	Local cIspb     := ""
	Local cCodJson  := ""
	Local cErro     := ""
	Local oJson     := Nil
	Local aBancos   := {}
	Local nI        := 0
	Local nHandle   := 0
	Local nTam      := 0

	// Caminho relativo ao RootPath (protheus_data). Nome SEMPRE em minusculo (Linux).
	cJsonPath := Lower(AllTrim(GetSrvProfString("StartPath", "\system\")))
	If Right(cJsonPath, 1) != "\"
		cJsonPath += "\"
	EndIf
	cJsonPath += "ispb.json"   // <-- confira se o nome bate com o arquivo no servidor

	// Normaliza o codigo: "1" -> "001", " 422 " -> "422"
	cCodBanco := StrZero(Val(AllTrim(cValToChar(cCodBanco))), 3)

	// 1. Verifica se o arquivo existe no servidor
	If !File(cJsonPath)
		MsgStop("Arquivo JSON nao encontrado em: " + cJsonPath, "Erro de Leitura")
		Return cIspb
	EndIf

	// 2. Le o conteudo inteiro do arquivo
	nHandle := FOpen(cJsonPath, 0)
	If nHandle == -1
		MsgStop("Nao foi possivel abrir o arquivo: " + cJsonPath + " (FError " + cValToChar(FError()) + ")", "Erro de IO")
		Return cIspb
	EndIf

	nTam := FSeek(nHandle, 0, 2)    // descobre o tamanho
	FSeek(nHandle, 0, 0)            // volta para o inicio
	cJsonData := Space(nTam)        // aloca o buffer
	FRead(nHandle, @cJsonData, nTam)
	FClose(nHandle)

	// 3. Deserializa. Envolve o array raiz num objeto para o JsonObject tratar.
	oJson := JsonObject():New()
	cErro := oJson:FromJson('{"lista":' + cJsonData + '}')
	If ValType(cErro) == "C" .And. !Empty(cErro)
		MsgStop("JSON invalido: " + cErro, "Erro de Leitura")
		Return cIspb
	EndIf

	aBancos := oJson["lista"]
	If ValType(aBancos) != "A"
		MsgStop("O JSON deve iniciar com uma lista [ ... ].", "Erro de Leitura")
		Return cIspb
	EndIf

	// DEBUG: mostra no console.log quantos itens e as chaves do 1o registro
	ConOut("FNispb - itens no JSON: " + cValToChar(Len(aBancos)))
	If Len(aBancos) > 0 .And. ValType(aBancos[1]) == "J"
		aEval(aBancos[1]:GetNames(), {|x| ConOut("FNispb - chave: [" + x + "] = " + cValToChar(aBancos[1][x])) })
	EndIf

	// 4. Procura o banco (nome da chave sem diferenciar maiusc/minusc)
	For nI := 1 To Len(aBancos)
		If ValType(aBancos[nI]) != "J"
			Loop
		EndIf

		cCodJson := AllTrim(cValToChar(JsonVal(aBancos[nI], {"codigo", "code", "compe", "cod", "numero_codigo", "codigo_compe", "banco"})))

		If !Empty(cCodJson) .And. StrZero(Val(cCodJson), 3) == cCodBanco
			cIspb := AllTrim(cValToChar(JsonVal(aBancos[nI], {"ispb"})))
			cIspb := IIf(Empty(cIspb), "", StrZero(Val(cIspb), 8))  // ISPB tem 8 digitos
			Exit
		EndIf
	Next nI

	ConOut("FNispb - banco [" + cCodBanco + "] ISPB retornado: [" + cIspb + "]")

	FreeObj(oJson)

Return cIspb

//-------------------------------------------------------------------
// Retorna o valor da 1a chave encontrada (ignora maiusc/minusc)
//-------------------------------------------------------------------
Static Function JsonVal(oItem, aChaves)
	Local aNomes := oItem:GetNames()
	Local nC     := 0
	Local nPos   := 0

	For nC := 1 To Len(aChaves)
		nPos := aScan(aNomes, {|x| Upper(AllTrim(x)) == Upper(aChaves[nC]) })
		If nPos > 0 .And. oItem[aNomes[nPos]] != Nil
			Return oItem[aNomes[nPos]]
		EndIf
	Next nC

Return ""

//-------------------------------------------------------------------
// Teste via SmartClient
//-------------------------------------------------------------------
User Function TesteIspb()
	Local cResultado := u_FNispb("422")

	If !Empty(cResultado)
		MsgInfo("O ISPB do banco 422 e: " + cResultado)
	Else
		MsgAlert("Banco nao localizado no JSON.")
	EndIf
Return

User Function fnChvPIX()
	Local _cReturn := StrZero(Val(SA2->A2_BANCO),3) + StrZero(Val(SA2->A2_AGENCIA),7) + StrZero(Val(SA2->A2_NUMCON),10)
Return(_cReturn)
