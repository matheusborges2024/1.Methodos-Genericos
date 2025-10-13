#Include "Protheus.ch"
#include "TOTVS.CH"
#INCLUDE "topconn.ch"
#include "rwmake.ch"
#INCLUDE "RPTDEF.CH"
#INCLUDE "FWPrintSetup.ch"

//** ROTINA UTILIZADA PARA A IMPRESSAO DO BOLETO DO BANCO
//** ITAU EM FORMATO GRAFICO COM CODIGO DE BARRAS



Class GeraBoleto
    data aRetorno 
    //Métodos
    Method New() CONSTRUCTOR
    Method ITAU()

EndClass



Method New() class GeraBoleto

return

Method ITAU(aTitulos,lJob,lArqUnic) Class GeraBoleto

    ::aRetorno:= u_BOL_341(aTitulos, lJob, lArqUnic)

return 






User Function BOL_341(_aImprBol, lJob, lArqUnic)

	Local cArea 		:= GetArea()
	Local aRet			:= {}

	Private cBanco		:= "341"

	If Len(_aImprBol)<=0
		Return()
	EndIf
	Processa({|lEnd| aRet := MontaRel(_aImprBol, lJob, lArqUnic)})

	//If aRet[1] .and. !lJob
	//	FWInputBox("E-mail do cliente para ser enviado o Boleto Bancário", aRet[4])
	//EndIf

	RestArea(cArea)

Return aRet


Static Function MontaRel(_aImprBol, lJob, lArqUnic)

	Local lRet			:= .F.
	Local _nItens
	Local _cEmailCl	:= ""
	Local cLinhaDig	:= ""
	Local nPerJuros	:= 0.02
	Private nX			:= 0
	Private cNroDoc	:=  " "
	Private aDadosEmp	:= {	AllTrim(SM0->M0_NOMECOM)														,;	//[1]Nome da Empresa
								AllTrim(SM0->M0_ENDCOB)															,;	//[2]Endereço
								AllTrim(SM0->M0_BAIRCOB)+", "+AllTrim(SM0->M0_CIDCOB)+", "+SM0->M0_ESTCOB		,;	//[3]Complemento
								"CEP: "+Subs(SM0->M0_CEPCOB,1,5)+"-"+Subs(SM0->M0_CEPCOB,6,3)					,;	//[4]CEP
								"PABX/FAX: "+SM0->M0_TEL														,;	//[5]Telefones
								"CNPJ: "+Subs(SM0->M0_CGC,1,2)+"."+Subs(SM0->M0_CGC,3,3)+"."+					;	//[6]
									Subs(SM0->M0_CGC,6,3)+"/"+Subs(SM0->M0_CGC,9,4)+"-"+						;	//[6]
									Subs(SM0->M0_CGC,13,2)														,;	//[6]CGC
								"I.E.: "+Subs(SM0->M0_INSC,1,3)+"."+Subs(SM0->M0_INSC,4,3)+"."+					;	//[7]
									Subs(SM0->M0_INSC,7,3)+"."+Subs(SM0->M0_INSC,10,3)							}	//[7]I.E

	Private aDadosTit		:= {}
	Private aDadosBanco	:= {}
	Private aDatSacado	:= {}
	Private aBolText		:= {	"Sujeito a protesto e negativacao"										,;
									"Cobrar multa de R$ "													,;
									"Apos o Vencimento cobrar juros de R$ "									,;
									""																		,;
									""																		}

	Private aBolTxtCp		:= {	""																		,;
									" apos o vencimento"													,;
									" ao dia"																,;
									""																		,;
									""																		}

	Private nI			:= 1
	Private aCB_RN_NN	:= {}
	Private nVlrAbat	:= 0

	Private cDir		:=getTempPath()
	Private nDevice		:= IMP_PDF	//IMP_SPOOL //IMP_PDF
	Private lAdjToLeg	:= .T.
	Private lDisabSet	:= .T. //lSetup
	Private lViewPDF	:= !lJob
	Private nQuant		:= 1	//Exencial para impressão em modo JOB

	If !lArqUnic

		Private oPrint

		If Len(_aImprBol) == 1
			cFilePrtx	:= "Boleto_NF_" +AllTrim(_aImprBol[1,3])+ "_Serie_" +AllTrim(_aImprBol[1,2])+ "_Parc_" +AllTrim(_aImprBol[1,4])+ "_dtImp_" + dToS(Date()) + "T" + Replace(Time(),":","")
		Else
			cFilePrtx	:= "Boletos_NF_" +AllTrim(_aImprBol[1,3])+ "_Serie_" +AllTrim(_aImprBol[1,2])+ "_dtImp_" + dToS(Date()) + "T" + Replace(Time(),":","")
		EndIf

		//oPrint := FWMSPrinter():New(cFilePrtx, nDevice, lAdjToLeg, cDir, lDisabSet, @oPrint, , , , , .F., lViewPDF, nQuant)
		oPrint := FWMSPrinter():New(cFilePrtx, nDevice, lAdjToLeg, cDir, lDisabSet, , @oPrint, , , , , lViewPDF, nQuant)
		oPrint:SetPortrait() // ou SetLandscape()
		//oPrint:StartPage()   // Inicia uma nova página
		oPrint:SetPaperSize(DMPAPER_A4)
		oPrint:SetMargin(10,10,10,10)

	EndIf

	ProcRegua(Len(_aImprBol))

	For _nItens := 1 to Len(_aImprBol)

		//Posiciona o SE1 (Titulos a Receber)
		DbSelectArea("SE1")
		DbSetOrder(1)
		If !DbSeek(_aImprBol[_nItens][1] + _aImprBol[_nItens][2] + _aImprBol[_nItens][3] + _aImprBol[_nItens][4] + _aImprBol[_nItens][5],.T.)
			If !lJob
				ApMsgStop("Titulo não encontrado : " + _aImprBol[_nItens][1] + _aImprBol[_nItens][2] + _aImprBol[_nItens][3] + _aImprBol[_nItens][4] + _aImprBol[_nItens][5] )
			EndIf
			Loop
		EndIf

		//VALIDA TITULO VENCIDO E SEM NUMERO BANCARIO
		If !lJob .and. (SE1->E1_VENCREA < dDataBase .Or. Empty(SE1->E1_NUMBCO))
			ApMsgStop("Titulo vencido ou sem Número Bancário")
			Loop
		EndIF

		//VALIDA TITULO COM PORTADOR DIFERENTE DO ITAU
		If !lJob .and. !SE1->E1_PORTADO $ cBanco
			ApMsgStop("Titulo não pertence ao Banco do Itau!")
			Loop
		EndIF

		//Posiciona o SA6 (Bancos)
		DbSelectArea("SA6")
		DbSetOrder(1)
		DbSeek(xFilial("SA6")+SE1->E1_PORTADO+SE1->E1_AGEDEP+SE1->E1_CONTA,.T.)

		//Posiciona na Arq de Parametros CNAB
		DbSelectArea("SEE")
		DbSetOrder(1)
		DbSeek(xFilial("SEE")+SE1->E1_PORTADO+SE1->E1_AGEDEP+SE1->E1_CONTA+_aImprBol[_nItens][6],.T.)

		//Posiciona o SA1 (Cliente)
		DbSelectArea("SA1")
		DbSetOrder(1)
		DbSeek(xFilial()+SE1->E1_CLIENTE+SE1->E1_LOJA,.T.)

		DbSelectArea("SE1")

		aDadosBanco	:= {	SA6->A6_COD																,;		// [1]Numero do Banco
							AllTrim(SA6->A6_NOME)													,;		// [2]Nome do Banco
							AllTrim(SA6->A6_AGENCIA)												,;		// [3]Agência
							SA6->A6_DVAGE															,;		// [4]Dígito da Agencia
							Substring(SA6->A6_NUMCON, 1,5)											,;		// [5]Conta Corrente
							Substring(SA6->A6_NUMCON, 7,1)											,;		// [6]Dígito da conta corrente
							allTrim(SEE->EE_VARCART)												,;		// [7]Codigo da Carteira
							Substring(SA6->A6_CODCED,1,7)											}		// [8]Codigo Cedente

		_cPessoa := IIF(Len(AllTrim(str(val(SA1->A1_CGC))))>11,"J","F")

		If Empty(SA1->A1_ENDCOB)
			aDatSacado := {	AllTrim(SA1->A1_NOME)																		,;	// [1]Razão Social
									AllTrim(SA1->A1_COD )+"-"+SA1->A1_LOJA												,;	// [2]Código
									AllTrim(SA1->A1_END )+" - "+AllTrim(SA1->A1_COMPLEM)+" - "+AllTrim(SA1->A1_BAIRRO)	,;	// [3]Endereço
									AllTrim(SA1->A1_MUN )																,;	// [4]Cidade
									SA1->A1_EST																			,;	// [5]Estado
									"CEP: "+Subs(SA1->A1_CEP,1,5)+"-"+Subs(SA1->A1_CEP,6,3)								,;	// [6]CEP
									SA1->A1_CGC																			,;	// [7]CGC
									_cPESSOA																			}	// [8]PESSOA
		Else
			aDatSacado := {	AllTrim(SA1->A1_NOME)																		,;	// [1]Razão Social
									AllTrim(SA1->A1_COD )+"-"+SA1->A1_LOJA												,;	// [2]Código
									AllTrim(SA1->A1_ENDCOB)+" - "+AllTrim(SA1->A1_BAIRROC)								,;	// [3]Endereço
									AllTrim(SA1->A1_MUNC)																,;	// [4]Cidade
									SA1->A1_ESTC																		,;	// [5]Estado
									"CEP: "+Subs(SA1->A1_CEPC,1,5)+"-"+Subs(SA1->A1_CEPC,6,3)							,;	// [6]CEP
									SA1->A1_CGC																			,;	// [7]CGC
									_cPESSOA																			}	// [8]PESSOA
		Endif

		nValTOri  := SE1->E1_SALDO
		dVencto   := SE1->E1_VENCTO
		dVencOri  := Iif(Empty(SE1->E1_VENCORI),dVencto,SE1->E1_VENCORI)

		If SE1->E1_PREFIXO == "PED"
			nValTAtu  := nValTOri
			aBolText	:= {	"Todas as informações deste bloqueto são de exclusiva responsabilidade do cedente."					,;
								"O pagamento deste boleto é a efetivação do seu pedido, de acordo com as condições comerciais."		,;
								"Pagando o boleto até a data de vencimento você garante as condições de entrega negociadas."		,;
								""																									,;
								""																		}
			aBolTxtCp  := {	"", "", "", "", ""}
		Else
			nValTAtu  := Round( ;
									nValTOri+	;		
									IF(dVencto>dVencOri.AND.(dVencto-dVencOri)>3	;
										,(dVencto - dVencOri) * ((nValTOri * nPerJuros)/30) + nValTOri * nPerJuros	;
										,IF(dVencto>dVencOri.AND.(dVencto-dVencOri)<=3 ;
											,(dVencto - dVencOri) * ((nValTOri * nPerJuros)/30)  ;
											,0)) ;
								,2)
		EndIf

		nVlrAbat  := SomaAbat(SE1->E1_PREFIXO,SE1->E1_NUM,SE1->E1_PARCELA,"R",1,,SE1->E1_CLIENTE,SE1->E1_LOJA)

		_cEmailCl := Alltrim(SA1->A1_EMAIL)	//IIF(EMPTY(SA1->A1_X_EMAIL),Alltrim(SA1->A1_EMAIL),Alltrim(SA1->A1_X_EMAIL))

		//Aqui defino parte do nosso numero. Sao 8 digitos para identificar o titulo.
		//Abaixo apenas uma sugestao.
		//cNroDoc	:= RIGHT( SE1->E1_NUMBCO,8)	//RIGHT( SE1->E1_IDCNAB,8)
		lGravaSee := .F.
		If Empty(SE1->E1_NUMBCO)
			cNroDoc	:= soma1(StrZero(Val(SEE->EE_FAXATU),8))
			lGravaSee := .T.
		Else
			cNroDoc := StrZero(Val(SE1->E1_NUMBCO),8)
		EndIf

		//Monta codigo de barras
		aCB_RN_NN	:= Ret_cB(Subs(aDadosBanco[1],1,3),"9",aDadosBanco[3],aDadosBanco[5],aDadosBanco[6],aDadosBanco[7],cNroDoc,nValTAtu,dVencto)
		cLinhaDig	:= aCB_RN_NN[2]
		cLinhaDig	:= replace(replace(cLinhaDig,' ',''), '.', '')
		_cDAC_NN	:= aCB_RN_NN[4]
		
		//GRAVA LINHA DIGITAVEL
		SE1->(RecLock("SE1",.F.))
		SE1->E1_CODBAR	:= aCB_RN_NN[1]
		SE1->E1_CODDIG := cLinhaDig
		SE1->(MsUnLock())

		If lGravaSee
			If Alltrim(cNroDoc) > StrZero(Val(SEE->EE_FAXFIM),8)
				Alert ("O nosso numero esta acima da faixa permitida. Favor rever a faixa no cadastro 'Parâmetros Bancos'.")
			Endif

			SEE->(RecLock("SEE",.F.))
			SEE->EE_FAXATU  := cNroDoc
			SEE->(MsUnLock())
		EndIf

		aDadosTit	:= {	AllTrim(SE1->E1_NUM)+AllTrim(SE1->E1_PARCELA)					,;	// [1] Número do título
							SE1->E1_EMISSAO													,;	// [2] Data da emissão do título
							dDataBase														,;	// [3] Data da emissão do boleto
							dVencto															,;	// [4] Data do vencimento
							nValTAtu														,;	// [5] Valor do título
							aCB_RN_NN[3]													,;	// [6] Nosso número (Ver fórmula para calculo)
							SE1->E1_PREFIXO													,;	// [7] Prefixo da NF
							SE1->E1_TIPO													,;	// [8] Tipo do Titulo
							SE1->E1_NUM														,;	// [9] Numero Titulo
							SE1->E1_PARCELA													,;	// [10] Parcela
							nValTOri														,;	// [11] Valor Original
							nPerJuros														,;	// [12] Percentual Juros
							_cDAC_NN														}	// [13] Dac Itau

		Impress(oPrint,aDadosEmp,aDadosTit,aDadosBanco,aDatSacado,aCB_RN_NN,aBolText,aBolTxtCp)
		nX := nX + 1

		IncProc()
		nI := nI + 1

	Next _nItens

	// Finaliza a página
	oPrint:EndPage()

	// Visualiza antes de imprimir
	If !Empty(aDadosTit) .and. !lArqUnic
		oPrint:Preview()
		lRet	:= .T.
	EndIf

Return {lRet, cDir, cFilePrtx, _cEmailCl}


Static Function Impress(oPrint,aDadosEmp,aDadosTit,aDadosBanco,aDatSacado,aCB_RN_NN,aBolText,aBolTxtCp)

	Local jj, nY, nI
	Local nColIni := -60

	//Parametros de TFont.New()  //1.Nome da Fonte (Windows)  //3.Tamanho em Pixels  //5.Bold (T/F)
	Local oFont10		:= TFont():New("Arial",9,10,.T.,.F.,5,.T.,5,.T.,.F.)
	Local oFont13n		:= TFont():New("Arial",9,13,.T.,.T.,5,.T.,5,.T.,.F.)
	Local oFont14n		:= TFont():New("Arial",9,14,.T.,.T.,5,.T.,5,.T.,.F.)
	Local oFont14c		:= TFont():New("Courier New",9,14,.T.,.T.,5,.T.,5,.T.,.F.)
	Local oFont16n		:= TFont():New("Arial",9,16,.T.,.T.,5,.T.,5,.T.,.F.)
	Local oFont20n		:= TFont():New("Arial",9,20,.T.,.T.,5,.T.,5,.T.,.F.)
	Local oFont24n		:= TFont():New("Arial",9,24,.T.,.T.,5,.T.,5,.T.,.F.)
	Local oFontLD		:= TFont():New("Times New Roman",8,22,.T.,.T.,5,.T.,5,.T.,.F.)

	oPrint:StartPage()   // Inicia uma nova página

	For nY := 1 to 2

		If nY == 1
			nRow1 := 50
		ElseIf nY == 2
			nRow1 := 1180
		EndIf

		For jj:=0 to 6 Step 1
			oPrint:Line (nRow1+126+jj,nColIni+100,nRow1+126+jj,nColIni+2300)
			oPrint:Line (nRow1+130,nColIni+430+jj,nRow1+050, nColIni+430+jj)
			oPrint:Line (nRow1+130,nColIni+650+jj,nRow1+050, nColIni+650+jj)
		Next jj

		cBitMap:= "ITAU_LOGO.BMP"
		oPrint:SayBitmap(nRow1+040,nColIni+100,cBitMap,080,080 )
		oPrint:Say  (nRow1+100,nColIni+188,"Banco Itaú SA"																								,oFont14n )	// [2]Nome do Banco
		oPrint:Say  (nRow1+110,nColIni+455,aDadosBanco[1]+"-7"																						,oFont24n )	// [1]Numero do Banco

		If nY == 1
			oPrint:Say  (nRow1+100,nColIni+1810,"Recibo do Pagador"																					,oFont20n)
		Else
			//Linha Digitavel do Codigo de Barras
			oPrint:Say  (nRow1+100,nColIni+700,aCB_RN_NN[2]																								,oFontLD)
		EndIf

		oPrint:Line (nRow1+0215,nColIni+0100,nRow1+0215,nColIni+2300 )
		oPrint:Line (nRow1+0295,nColIni+0100,nRow1+0295,nColIni+2300 )
		oPrint:Line (nRow1+0365,nColIni+0100,nRow1+0365,nColIni+2300 )
		oPrint:Line (nRow1+0430,nColIni+0100,nRow1+0430,nColIni+2300 )
		oPrint:Line (nRow1+0495,nColIni+0100,nRow1+0495,nColIni+2300 )

		oPrint:Line (nRow1+0130,nColIni+1800,nRow1+0755,nColIni+1800)
		oPrint:Line (nRow1+0365,nColIni+0500,nRow1+0495,nColIni+0500)
		oPrint:Line (nRow1+0430,nColIni+0750,nRow1+0495,nColIni+0750)
		oPrint:Line (nRow1+0365,nColIni+1000,nRow1+0495,nColIni+1000)
		oPrint:Line (nRow1+0365,nColIni+1300,nRow1+0430,nColIni+1300)
		oPrint:Line (nRow1+0365,nColIni+1480,nRow1+0495,nColIni+1480)
		
		oPrint:Say  (nRow1+160,nColIni+100 ,"Local de Pagamento:"																					,oFont10)
		oPrint:Say  (nRow1+200,nColIni+100 ,"PREFERENCIALMENTE NO ITAU ATE O VENCIMENTO. APOS O VENCIMENTO, SOMENTE NO ITAU."	,oFont14n)

		oPrint:Say  (nRow1+160,nColIni+1810,"Vencimento:"																								,oFont10)
		cString	:= StrZero(Day(aDadosTit[4]),2) +"/"+ StrZero(Month(aDadosTit[4]),2) +"/"+ Right(Str(Year(aDadosTit[4])),4)
		nCol := 1850+(375-(len(cString)*22))
		oPrint:Say  (nRow1+200,nColIni+nCol,cString																										,oFont16n)

		oPrint:Say  (nRow1+240,nColIni+100 ,"Beneficiário:"																							,oFont10)
		oPrint:Say  (nRow1+280,nColIni+100 ,aDadosEmp[1]+" - "+aDadosEmp[6]																		,oFont14n)	//Nome + CNPJ

		oPrint:Say  (nRow1+240,nColIni+1810,"Agência/Código Beneficiário:"																		,oFont10)
		cString := RIGHT( SPACE(15) + Alltrim(aDadosBanco[3]+"/"+aDadosBanco[5]+"-"+aDadosBanco[6]), 15)
		nCol := 1832+(440-(len(cString)*22))
		oPrint:Say  (nRow1+280,nColIni+nCol,cString																										,oFont14c)

		oPrint:Say  (nRow1+320,nColIni+100 ,"Endereço Beneficiário / Beneficiário Avalista:"												,oFont10)
		oPrint:Say  (nRow1+355,nColIni+100 ,aDadosEmp[2] + " - " + aDadosEmp[3] + " - " + aDadosEmp[4]									,oFont14n)	//Nome + CNPJ

		oPrint:Say  (nRow1+320,nColIni+1810,"Nosso Número:"																							,oFont10)
		_cDAC_NN := aDadosTit[13]
		cTitulo := RIGHT( SPACE(15) + aDadosBanco[7] + "/" + StrZero(Val(aDadosTit[6]),08) + "-" + _cDAC_NN,15)
		nCol := 1832+(440-(len(cTitulo)*22))
		oPrint:Say  (nRow1+355,nColIni+nCol,cTitulo																										,oFont14c)

		oPrint:Say  (nRow1+390,nColIni+100 ,"Data do Documento:"																						,oFont10)
		oPrint:Say  (nRow1+420,nColIni+100, StrZero(Day(aDadosTit[2]),2) +"/"+ StrZero(Month(aDadosTit[2]),2) +"/"+ Right(Str(Year(aDadosTit[2])),4)		,oFont13n)

		oPrint:Say  (nRow1+390,nColIni+505 ,"Nro.Documento:"																							,oFont10)
		oPrint:Say  (nRow1+420,nColIni+605 ,aDadosTit[7] + "  " + aDadosTit[9] + "  " + aDadosTit[10]									,oFont13n)	//Prefixo +Numero+Parcela

		oPrint:Say  (nRow1+390,nColIni+1005,"Espécie Doc.:"																							,oFont10)
		oPrint:Say  (nRow1+420,nColIni+1050,"DM"																											,oFont13n)	//Tipo do Titulo

		oPrint:Say  (nRow1+390,nColIni+1305,"Aceite:"																									,oFont10)
		oPrint:Say  (nRow1+420,nColIni+1400,"N"																											,oFont13n)

		oPrint:Say  (nRow1+390,nColIni+1485,"Data do Processamento:"																				,oFont10)
		oPrint:Say  (nRow1+420,nColIni+1550,StrZero(Day(aDadosTit[3]),2) +"/"+ StrZero(Month(aDadosTit[3]),2) +"/"+ Right(Str(Year(aDadosTit[3])),4)		,oFont13n) // Data impressao

		oPrint:Say  (nRow1+390,nColIni+1810,"(=) Valor do Documento:"																				,oFont10)
		cString := RIGHT( SPACE(15) + Alltrim(Transform(aDadosTit[5],"@E 99,999,999.99")),15)
		nCol := 1832+(440-(len(cString)*22))
		oPrint:Say  (nRow1+420,nColIni+nCol,cString																										,oFont14c)

		oPrint:Say  (nRow1+455,nColIni+100 ,"Uso do Banco:"																							,oFont10)

		oPrint:Say  (nRow1+455,nColIni+505 ,"Carteira:"																									,oFont10)
		oPrint:Say  (nRow1+485,nColIni+555 ,aDadosBanco[7]																								,oFont13n)

		oPrint:Say  (nRow1+455,nColIni+755 ,"Espécie:"																									,oFont10)
		oPrint:Say  (nRow1+485,nColIni+805 ,"R$"																											,oFont13n)

		oPrint:Say  (nRow1+455,nColIni+1005,"Quantidade:"																								,oFont10)

		oPrint:Say  (nRow1+455,nColIni+1485,"Valor:"																										,oFont10)

		oPrint:Say  (nRow1+455,nColIni+1810,"(-) Desconto / Abatimento:"																			,oFont10)

		oPrint:Say  (nRow1+520,nColIni+100 ,"Instruções: (Todas informações deste bloqueto são de exclusiva responsabilidade do Beneficiário)"		,oFont10)

		cString1 := ""
		cString2 := ""
		If aDadosTit[7] <> "PED"
			cString1 := ALLTRIM(TRANSFORM(((aDadosTit[11])*aDadosTit[12]),PESQPICT("SE1","E1_VALOR")))
			cString2 := ALLTRIM(TRANSFORM(((aDadosTit[11])*aDadosTit[12])/30,PESQPICT("SE1","E1_VALOR")))
		EndIf
		oPrint:Say  (nRow1+570,nColIni+120 ,aBolText[1] + aBolTxtCp[1]																				,oFont14n)
		oPrint:Say  (nRow1+610,nColIni+120 ,aBolText[2] + cString1 + aBolTxtCp[2] 																,oFont14n)
		oPrint:Say  (nRow1+650,nColIni+120 ,aBolText[3] + cString2 + aBolTxtCp[3]																,oFont14n)
		oPrint:Say  (nRow1+690,nColIni+120 ,aBolText[4] + aBolTxtCp[4]																				,oFont14n)
		oPrint:Say  (nRow1+730,nColIni+120 ,aBolText[5] + aBolTxtCp[5]																				,oFont14n)

		oPrint:Say  (nRow1+520,nColIni+1810,"(-) Outras Deduções:"																					,oFont10)

		oPrint:Line (nRow1+560,nColIni+1800,nRow1+560,nColIni+2300 )
		oPrint:Line (nRow1+625,nColIni+1800,nRow1+625,nColIni+2300 )
		oPrint:Line (nRow1+690,nColIni+1800,nRow1+690,nColIni+2300 )
		oPrint:Line (nRow1+755,nColIni+100 ,nRow1+755,nColIni+2300 )

		oPrint:Say  (nRow1+585,nColIni+1810,"(+) Mora/Multa:"																							,oFont10)
		oPrint:Say  (nRow1+650,nColIni+1810,"(+) Outros Acréscimos:"																					,oFont10)
		oPrint:Say  (nRow1+715,nColIni+1810,"(=) Valor Cobrado:"																						,oFont10)

		If aDatSacado[8] = "J"
			cDocto := "CNPJ: "+TRANSFORM(aDatSacado[7],"@R 99.999.999/9999-99")
		Else
			cDocto := "CPF: "+TRANSFORM(aDatSacado[7],"@R 999.999.999-99")
		EndIf

		oPrint:Say  (nRow1+780,nColIni+100 ,"Sacado:"																									,oFont10)
		oPrint:Say  (nRow1+785,nColIni+250 ,aDatSacado[1]+" ("+aDatSacado[2]+")" + "  -  "+cDocto											,oFont13n)	// Razão Social + Docto
		oPrint:Say  (nRow1+820,nColIni+0250 ,aDatSacado[3]																								,oFont13n)	// Endereço
		oPrint:Say  (nRow1+855,nColIni+0250 ,aDatSacado[6]+"    "+aDatSacado[4]+" - "+aDatSacado[5]										,oFont13n)	// CEP+Cidade+Estado

		oPrint:Say  (nRow1+900,nColIni+0100 ,"Sacador/Avalista:"																						,oFont10)
		oPrint:Say  (nRow1+900,nColIni+1720 ,"Código de Baixa"																						,oFont10)

		oPrint:Say  (nRow1+900,nColIni+1950 ,cTitulo																										,oFont13n)

		For jj:=0 to 6 Step 1
			oPrint:Line (nRow1+915+jj,nColIni+100 ,nRow1+915+jj,nColIni+2300 )
		Next jj

		oPrint:Say  (nRow1+955,nColIni+1900 ,"Autenticação Mecânica"																				,oFont10)

		If nY == 1
			//Pontilhado separador
			For nI := 100 to 2280 step 25
				oPrint:Line(nRow1+1120, nColIni+nI,nRow1+1120, nColIni+nI+15)
			Next nI
		Else
			//Codigo de Barras
			oPrint:FwMSBAR("INT25" ,50, 2.5,aCB_RN_NN[1],oPrint,.F. ,Nil ,Nil ,0.017 ,1 ,Nil ,Nil,"A" ,.F. )
		EndIf

	Next nY

	// Finaliza a página
	oPrint:EndPage()

Return Nil


Static Function Modulo10(cData)

	Private L,D,P := 0
	Private B     := .F.

	L := Len(cData)
	B := .T.
	D := 0
	While L > 0
		P := Val(SubStr(cData, L, 1))
		If (B)
			P := P * 2
			If P > 9
				P := P - 9
			End
		End
		D := D + P
		L := L - 1
		B := !B
	End

	D := 10 - (Mod(D,10))
	If D = 10
		D := 0
	End

Return(D)


Static Function Modulo11(cData)

	Private L, D, P := 0

	L := Len(cdata)
	D := 0
	P := 1

	While L > 0
		P := P + 1
		D := D + (Val(SubStr(cData, L, 1)) * P)
		If P = 9
			P := 1
		End
		L := L - 1
	End

	D := 11 - (mod(D,11))
	If (D == 0 .Or. D == 1 .Or. D == 10 .Or. D == 11)
		D := 1
	End
	
Return(D)


Static Function Ret_CB(cBanco,cMd,cAgencia,cConta,cDacCC,cCart,cNroDoc,nValor,dVencto)
	
	Local nFator := 0

	Private cValorFinal := Replicate("0",10 - Len(allTrim(STR(nValor*100)))) + allTrim(STR(nValor*100))
	Private nDvnn		:= 0
	Private nDvcb		:= 0
	Private nDv			:= 0
	Private cNN			:= ''
	Private cRN			:= ''
	Private cCB			:= ''
	Private cS			:= ''
	Private cFator    := ""		//strzero(dVencto - ctod("07/10/97"),4)

	nFator := dVencto - ctod("07/10/97")
	If nFator > 9999
		nFator := nFator - 9000
	Endif

	cFator := subs(strzero(nFator,4),1,4)
	
	//-----------------------------
	// Definicao do NOSSO NUMERO
	// ----------------------------
	cNN   := StrZero(Val(cNroDoc),8)
	_cDC := AllTrim(Str(modulo10(cAgencia + cConta + cCart + cNN)))

	//----------------------------------
	//	 Definicao da LINHA DIGITAVEL 
	//----------------------------------
	//-------- Definicao da LINHA DIGITAVEL (Representacao Numerica)
	//	Campo 1			Campo 2			Campo 3		 Campo 4	Campo 5		
	//	AAABC.CCCDX		DDDDD.DEFFFY	FFFFF.FFFFZ	 K			HHHHIIIIIIIIII

	// 	CAMPO 1:
	//	AAA	 		= Codigo do banco na Camara de Compensacao 341
	//	  B  		= Codigo da moeda, sempre 9
	//	CCCC 		= Agencia
	//	  D  		= Primeiro digito na conta corrente
	//	  X 		= DAC que amarra o campo, calculado pelo Modulo 10 da String do campo
	_cBlc1 		:= cBanco + "9" + cCart + Substr(cNN,1,2)
	_cDgBlc1 	:= AllTrim(Str(modulo10(_cBlc1)))

	// 	CAMPO 2:
	//	DDDDDD 		= Restante do numero da conta corrente
	//  E			= Digito da Cobranca
	//  FFF			= Tres primeiros digitos do Nosso Numero
	//	  Y 		= DAC que amarra o campo, calculado pelo Modulo 10 da String do campo
	_cBlc2 		:= Substr(cNN,3,6) + _cDC + Substr(cAgencia,1,3)
	_cDgBlc2 	:= AllTrim(Str(modulo10(_cBlc2)))

	// 	CAMPO 3:
	//	FFFFFFFFFF 	= Restante do nosso numero 
	//	  Z 		= DAC que amarra o campo, calculado pelo Modulo 10 da String do campo
	_cBlc3 		:= Substr(cAgencia,4,1) + cConta + cDacCC + "000"
	_cDgBlc3 	:= AllTrim(Str(modulo10(_cBlc3)))

	// 	CAMPO 4:    
	//  K			= Digitao do codigo de barras
	K:= cBanco + "9" + cFator +  cValorFinal + cCart + cNN + _cDC + cAgencia + cConta + cDACCC + "000"
	_cDg_K := AllTrim(Str(modulo11(K)))

	// 	CAMPO 5:
	//	HHHH   		= Fator de Vencimento
	//	IIIIIIIIII 	= Valor do Titulo
	_cBlc4 		:= cFator + cValorFinal
	//MessageBox(_cBlc4)

	cRN := Substr(_cBlc1 + _cDgBlc1,1,5)+"."+Substr(_cBlc1 + _cDgBlc1,6,5) + "  "
	cRN := cRN + Substr(_cBlc2 + _cDgBlc2,1,5) + "." +Substr(_cBlc2 + _cDgBlc2,6,6) + "  "
	cRN := cRN + Substr(_cBlc3 + _cDgBlc3,1,5) + "." +Substr(_cBlc3 + _cDgBlc3,6,6) + "  "
	cRN := cRN + _cDg_K + "  " + _cBlc4

	//----------------------------------
	//	 Definicao do CODIGO DE BARRAS
	//----------------------------------
	cCB		:= cBanco + "9" + _cDg_K + cFator +  cValorFinal + cCart + cNN + _cDC + cAgencia + cConta + cDACCC + "000"

Return({cCB,cRN,cNN,_cDC})
