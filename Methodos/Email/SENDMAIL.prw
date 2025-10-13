#include 'protheus.ch'
#include 'totvs.ch'
#INCLUDE 'TOPCONN.CH'
#INCLUDE "TBICONN.ch"

Class SENDMAIL
	DATA HTML
	DATA cMailEnv
	DATA cAssunto
	DATA oMessage
	DATA oLog
	DATA oServer
	data cErro
	data lGeralog
	Data cRootPath
    Data lEmail
	Method New() Constructor
	Method HTML()
	Method Send()
	Method ConnectionBuild()
	Method CloseConnection()
	Method BuildFile()
	Method XPUTSX6()
EndClass

Method New(cMailEnv,cAssunto,lEmail) class SENDMAIL
	::XPUTSX6()
	local cPasta 	:= "C:\logBoleto\"
	local cArquivo  := "LogBoleto.txt"
	local lHora		:= .t.
	::lGeralog 		:= GETMV("XX_GERALOG")
	::cMailEnv 		:= cMailEnv
	::cAssunto 		:= cAssunto
	::cRootPath 	:=  "/Arquivos_Email/"
    ::lEmail        :=  lEmail
	::oLog	   		:= GERALOG():New(cPasta, cArquivo, lHora,.f.,.f.)
	::HTML()

return

Method Send(cHtml,aArq) Class SENDMAIL
	local nI := 0
	local nErro := 0
	local lret := .T.
	local cFile := ""

	self:oMessage := TMailMessage():New()
	self:oMessage:Clear()
	self:oMessage:cFrom := ALLTRIM(GETMV("XX_MAIL"))
	self:oMessage:cTo 	:= SELF:cMailEnv
	self:oMessage:cSubject := EncodeUTF8(::cAssunto)
	self:oMessage:cBody := EncodeUTF8(cHtml)
	IF aArq <> nil .AND. Len(aArq)
		for nI := 1 to len(aArq)
			cFile:= self:BuildFile(aArq[nI])
			self:oMessage:AttachFile(cFile)
		Next
	endif
	self:oMessage:AddCustomHeader( "Content-Type", "text/calendar" )
	//Envia e-mail
	nErro := self:oMessage:Send( self:oServer )
	//Verifica se o e-mail foi enviado
	if nErro <> 0
		::oLog:AddText("Envio com Erro para: "+ SELF:cMailEnv)
		::oLog:AddText("Erro: "+ self:oServer:GetErrorString( nErro ))
		::oLog:Finish()
		conout( "ERROR:" + self:oServer:GetErrorString( nErro ) )
		::cErro :=  self:oServer:GetErrorString( nErro )
		self:oServer:SMTPDisconnect()
		nErro ++
		lret:= .f.
	else
		::oLog:AddText("Envio com sucesso para: "+ SELF:cMailEnv)
		::oLog:Finish()
	endif
	conout( 'Desconectando do SMTP' )
Return lret


Method CloseConnection() class SENDMAIL 
	self:oServer:SMTPDisconnect()
return


Method ConnectionBuild() class SENDMAIL
	LOCAL lRet := .T.
	self:oServer := TMailManager():New()
	self:oServer:SetUseSSL(GETMV("XX_SSL"))
	self:oServer:SetUseTLS( GETMV("XX_TSL"))
	self:oServer:Init( ALLTRIM(GETMV("XX_SMTP")), ALLTRIM(GETMV("XX_ACCONT")) , ALLTRIM(GETMV("XX_MAIL")), ALLTRIM(GETMV("XX_PASS")), 993, 587 )
	self:oServer:SetSmtpTimeOut( 300 )
	//Verifica conexão SMTP
	conout( 'Conectando do SMTP' )
	nErro := self:oServer:SmtpConnect()
	If nErro <> 0
		::oLog:AddText("Envio com Erro para: "+ SELF:cMailEnv)
		::oLog:AddText("Erro: "+ self:oServer:GetErrorString( nErro ))
		conout( "ERROR:" + self:oServer:GetErrorString( nErro ) )
		::cErro :=  self:oServer:GetErrorString( nErro )
		self:oServer:SMTPDisconnect()
		lRet :=.F.
	Endif

	//Verifica autenticação
	nErro := self:oServer:SmtpAuth( ALLTRIM(GETMV("XX_MAIL")) ,ALLTRIM(GETMV("XX_PASS")) )
	If nErro <> 0
		::oLog:AddText("Envio com Erro para: "+ SELF:cMailEnv)
		::oLog:AddText("Erro: "+ self:oServer:GetErrorString( nErro ))
		::cErro :=  self:oServer:GetErrorString( nErro )
		conout( "ERROR:" + self:oServer:GetErrorString( nErro ) )
		self:oServer:SMTPDisconnect()
		lRet := .F.
	Endif
return lRet


Method HTML() CLASS SENDMAIL

	::Html := "<!DOCTYPE html>" + CRLF
	::Html += '<html lang="pt-br">' + CRLF
	::Html += '<head>' + CRLF
	::Html += '    <meta charset="UTF-8">' + CRLF
	::Html += '    <meta name="viewport" content="width=device-width, initial-scale=1.0">' + CRLF
	::Html += '    <meta http-equiv="X-UA-Compatible" content="IE=edge">' + CRLF
	::Html += '    <title>{{titulo}}</title>' + CRLF
	::Html += '    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' + CRLF
	::Html += '    </head>' + CRLF
	::Html += '<body style="margin: 0; padding: 0; width: 100%; background-color: #f2f5f8; font-family: Arial, Helvetica, sans-serif;">' + CRLF
	::Html += '' + CRLF
	::Html += '    <div style="display:none; font-size:1px; color:#ffffff; line-height:1px; max-height:0px; max-width:0px; opacity:0; overflow:hidden;">' + CRLF
	::Html += '        Envio para o Cliente Dentro da VV' + CRLF
	::Html += '    </div>' + CRLF
	::Html += '' + CRLF
	::Html += '    <table role="presentation" align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse; background-color: #f2f5f8;">' + CRLF
	::Html += '        <tr>' + CRLF
	::Html += '            <td align="center" style="padding: 20px;">' + CRLF
	::Html += '                <table role="presentation" align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 600px; border-collapse: collapse; background-color: #ffffff; border-radius: 12px; border: 1px solid #dee2e6; overflow: hidden;">' + CRLF
	::Html += '                    ' + CRLF
	::Html += '                    <tr>' + CRLF
	::Html += '                        <td align="center" style="padding: 40px 0 30px 0;">' + CRLF
	::Html += '                            <img src="{{logo}}" alt="Logo da Empresa" width="150" style="display: block; border: 0;">' + CRLF
	::Html += '                        </td>' + CRLF
	::Html += '                    </tr>' + CRLF
	::Html += '' + CRLF
	::Html += '                    <tr>' + CRLF
	::Html += '                        <td style="padding: 0 40px; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 1.6; color: #343a40;">' + CRLF
	::Html += '                            <h1 style="font-size: 26px; margin: 0 0 15px; color: #212529; text-align: center; font-weight: 600;">{{titulo}}</h1>' + CRLF
	::Html += '                            <p style="margin: 0 0 25px; text-align: center; font-size: 18px; color: #6c757d;">' + CRLF
	::Html += '                                {{resumo}}' + CRLF
	::Html += '                            </p>' + CRLF
	::Html += '                                {{mensagem}}' + CRLF
	::Html += '                        </td>' + CRLF
	::Html += '                    </tr>' + CRLF
	::Html += '' + CRLF
	::Html += '                    <tr>' + CRLF
	::Html += '                        <td style="padding: 40px 30px 30px; background-color: #f8f9fa; border-top: 1px solid #dee2e6; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.5; color: #6c757d; text-align: center;">' + CRLF
	::Html += '                            <p style="margin: 0 0 15px;">{{dados.empresa}}</p>' + CRLF
	::Html += '                            ' + CRLF
	::Html += '                            <div style="font-size: 12px; color: #adb5bd;">' + CRLF
	::Html += '                                <p style="margin: 0;">Este é um e-mail automático, por favor, não responda.</p>' + CRLF
	::Html += '                                <p style="margin: 5px 0 0; word-break: break-all;"><a href="mailto:{{email.empresa}}" style="color: #adb5bd; text-decoration: underline;">{{email.empresa}}</a></p>' + CRLF
	::Html += '                            </div>' + CRLF
	::Html += '                        </td>' + CRLF
	::Html += '                    </tr>' + CRLF
	::Html += '' + CRLF
	::Html += '                </table>' + CRLF
	::Html += '            </td>' + CRLF
	::Html += '        </tr>' + CRLF
	::Html += '    </table>' + CRLF
	::Html += '' + CRLF
	::Html += '</body>' + CRLF
	::Html += '</html>' + CRLF

Return

method XPUTSX6() CLASS SENDMAIL

	local lRet := .f.
	Local aPars := {}

	Local nAtual        := 0
	Local aArea          := GetArea()
	Local aAreaX6        := SX6->(GetArea())
	Default aPars        := {}

	aAdd(aPars, {"XX_MAIL"   , "C", "E-mail de envio"          			,""})
	aAdd(aPars, {"XX_PASS" 	 , "C", "Senha do e-mail"   	  			,""})
	aAdd(aPars, {"XX_SMTP"   , "C", "SMPT DO E-MAIL"     				,""})
	aAdd(aPars, {"XX_ACCONT" , "C", "Conta de usuario no Servidor"     	,""})
	aAdd(aPars, {"XX_SSL"    , "L", "SSL"     							,""})
	aAdd(aPars, {"XX_TSL"	 , "L", "TSL"      							,""})
	aAdd(aPars, {"XX_GERALOG", "L", "GERALOG DE ARQUIVO"				,""})
	DbSelectArea("SX6")
	SX6->(DbGoTop())

	For nAtual := 1 To Len(aPars)
		If !(SX6->(DbSeek(xFilial("SX6")+aPars[nAtual][1])))
			RecLock("SX6",.T.)

			//Geral
			X6_VAR        :=    aPars[nAtual][1]
			X6_TIPO       :=    aPars[nAtual][2]
			X6_PROPRI     :=    "U"
			//Descricao
			X6_DESCRIC    :=    aPars[nAtual][3]
			X6_DSCSPA    :=    aPars[nAtual][3]
			X6_DSCENG    :=    aPars[nAtual][3]
			//Conteudo
			X6_CONTEUD    :=    aPars[nAtual][4]
			X6_CONTSPA    :=    aPars[nAtual][4]
			X6_CONTENG    :=    aPars[nAtual][4]
			SX6->(MsUnlock())
			lRet := .t.
		EndIf
	Next

	RestArea(aAreaX6)
	RestArea(aArea)

Return


Method BuildFile(cFileName) class SENDMAIL
	local nAtual := 0
	local cChave1 := "\"
	Local cChave2 := "/"
	local nAux := 0

	If !ExistDir(::cRootPath)
		MakeDir(::cRootPath)
	EndIf

	For nAtual := 1 To Len(cFileName)
		//Se a posição atual for igual ao caracter procurado, incrementa o valor
		If SubStr(cFileName, nAtual, 1) == cChave1
			nAux:= nAtual
		EndIf
	Next
	if nAux == 0
		For nAtual := 1 To Len(cFileName)
			//Se a posição atual for igual ao caracter procurado, incrementa o valor
			If SubStr(cFileName, nAtual, 1) == cChave2
				nAux:= nAtual
			EndIf
		Next
	endif
	if nAux > 0
		_CopyFile(cFileName,::cRootPath+SubStr(cFileName,nAux+1))
		return ::cRootPath+SubStr(cFileName,nAux+1)
	endif

return
