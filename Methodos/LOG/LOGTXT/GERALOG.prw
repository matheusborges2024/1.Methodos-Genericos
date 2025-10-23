//Bibliotecas
#Include "TOTVS.ch"
#include "fileio.ch"
 
Class GERALOG
    //Atributos
    Data cDirectory
    Data cFileName
    Data lShowTime
    Data oFWriter
    Data lOpenFile
    DATA nHandle
 
    //Métodos
    Method New() CONSTRUCTOR
    Method AddText()
    Method Finish()
EndClass
 
Method New(cDir, cFile, lShow, lOpen) Class GERALOG
    Default cDir  := GetTempPath()
    Default cFile := "log_" + dToS(Date()) + "_" + StrTran(Time(), ":", "-") + ".txt"
    Default lShow := .T.
    Default lOpen := .T.
 
    //Se a pasta não existir, cria ela
    If ! ExistDir(cDir)
        MakeDir(cDir)
    EndIf
    //Define os atributos
    ::cDirectory := cDir
    ::cFileName  := cFile 
    ::lShowTime  := lShow
    ::lOpenFile  := lOpen

    if !file(::cDirectory +  ::cFileName)
        ::nHandle:=FCREATE(::cDirectory +  ::cFileName) 
        FClose(::nHandle)
    endif
    //Cria o arquivo de logs
    ::nHandle := fopen(::cDirectory +  ::cFileName, FO_READWRITE + FO_SHARED )
    FSeek(::nHandle, 0, FS_END)
    FWrite(::nHandle,"/=================================================================/")
    FWrite(::nHandle,CRLF)
    FWrite(::nHandle,"Código do Usuário: " + RetCodUsr() )
    FWrite(::nHandle,CRLF)
    FWrite(::nHandle,"Nome do Usuário:   " + UsrRetName(RetCodUsr()) )
    FWrite(::nHandle,CRLF)
    FWrite(::nHandle,"Função (FunName):  " + FunName() )
    FWrite(::nHandle,CRLF)
    FWrite(::nHandle,"Ambiente:          " + GetEnvServer() )
    FWrite(::nHandle,CRLF)
    //EndIf
Return Self
 
Method AddText(cText) Class GERALOG
    Default cText := ""
 
    //Se for mostrar a hora, adiciona ela a esquerda
    If ::lShowTime
        cText := "[" + Time() + "] " + cText
    EndIf
    FWrite(::nHandle, alltrim(cText) ) // Insere texto no arquivo
    FWrite(::nHandle, CRLF) // Insere texto no arquivo

Return
 
Method Finish() Class GERALOG

    fclose(::nHandle) 
    //Se não for via job/webservice e tiver definido para abrir o arquivo
    If ! IsBlind() .And. ::lOpenFile
        ShellExecute("OPEN", ::cFileName, "", ::cDirectory, 1)
    EndIf
Return
