#include 'protheus.ch'
#include 'totvs.ch'
#INCLUDE 'TOPCONN.CH'
#INCLUDE "TBICONN.ch"

Class LOGAPI
	Method New() Constructor
    Method CreateTable()
    Method SEND()
EndClass


Method New() class LOGAPI
    ::CreateTable()

return


Method SEND(PROGRAMA,URL,ENDPOINT,METODO,REQUEST_BODY,RESPONSE_BODY,STATUS_CODE,TEMPO_RESPOSTA,USUARIO,AMBIENTE,OBSERVACAO) CLASS LOGAPI
    local cInsert := ""
    local lRet := .t. 
    // ========= Bloco de Tratameto de Dados Antes do INsert=========
        do Case
            Case Empty(PROGRAMA) .or. PROGRAMA == nil
                lRet := .F.
            Case Empty(URL) .or. URL == nil
                lRet := .F.
            Case Empty(ENDPOINT) .or. ENDPOINT == nil
                lRet := .F.            
            Case Empty(AMBIENTE) .or. AMBIENTE == nil
                lRet := .F.
            Case Empty(METODO) .or. METODO == nil
                lRet := .F.             
        ENDCASE      
    // ========= Bloco de Tratameto de Dados Antes do INsert=========
    if lRet 
        cInsert += "INSERT INTO API_LOG ("+CRLF
        cInsert += "    PROGRAMA,"+CRLF 
        cInsert += "    ORIGEM,"+CRLF 
        cInsert += "    ENDPOINT,"+CRLF 
        cInsert += "    METODO,"+CRLF 
        cInsert += "    REQUEST_BODY,"+CRLF 
        cInsert += "    RESPONSE_BODY,"+CRLF 
        cInsert += "    STATUS_CODE,"+CRLF 
        cInsert += "    TEMPO_RESPOSTA,"+CRLF 
        cInsert += "    USUARIO,"+CRLF 
        cInsert += "    AMBIENTE,"+CRLF 
        cInsert += "    OBSERVACAO"+CRLF
        cInsert += ") VALUES ("+CRLF
        cInsert += "'"+PROGRAMA+"',"+CRLF
        cInsert += "'"+URL+"',"+CRLF
        cInsert += "'"+ENDPOINT+"',"+CRLF
        cInsert += "'"+METODO+"',"+CRLF
        cInsert += "'"+REQUEST_BODY+"',"+CRLF
        cInsert += "'"+RESPONSE_BODY+"',"+CRLF
        cInsert += ""+cValToChar(STATUS_CODE)+","+CRLF
        cInsert += ""+cValToChar(TEMPO_RESPOSTA)+","+CRLF
        cInsert += "'"+USUARIO+"',"+CRLF
        cInsert += "'"+AMBIENTE+"',"+CRLF
        cInsert += "'"+OBSERVACAO+"'"+CRLF
        cInsert += ");"

        
        If TCSqlExec(cInsert) < 0
	        conout("Não foi possivel inserir Dados Dentro da Tabela LOGAPI, Motivo: "+TCSQLError() )
            lRet:=.F.
	    endif
    endif
Return lRet

Method CreateTable() Class LOGAPI

	local cScript := ""
    local cQuery  := ""
    local cTable  := "API_LOG"
    local nCount  := 0
	DO CASE
	CASE TCGetDB() == "MSSQL"
		cQuery := "SELECT *" + CRLF
		cQuery += "FROM" + CRLF
		cQuery += "     INFORMATION_SCHEMA.TABLES" + CRLF
		cQuery += "WHERE TABLE_NAME = '" + cTable + "'" + CRLF

	CASE TCGetDB() == "ORACLE"
		cQuery := "SELECT table_name" + CRLF
		cQuery += "FROM user_tables" + CRLF
		cQuery += "WHERE table_name = '" + Upper(cTable) + "'" + CRLF

	CASE TCGetDB() == "POSTGRES"
		cQuery := "SELECT to_regclass('public." + cTable + "')" + CRLF
	CASE TCGetDB() == "MYSQL"
		cQuery := "SELECT TABLE_NAME" + CRLF
		cQuery += "FROM INFORMATION_SCHEMA.TABLES" + CRLF
		cQuery += "WHERE TABLE_SCHEMA = DATABASE()" + CRLF
		cQuery += "  AND TABLE_NAME = '" + cTable + "'" + CRLF
	ENDCASE

	IF SELECT ("LOG") >0
		LOG->(DBCLOSEAREA())
	ENDIF
	TcQuery cQuery New Alias "LOG"
	count to nCount
	IF SELECT ("LOG") >0
		LOG->(DBCLOSEAREA())
	ENDIF

    if nCount <= 0
	    cScript := "CREATE TABLE API_LOG ("+CRLF
	    cScript += "ID_LOG          BIGINT IDENTITY(1,1) PRIMARY KEY,"+CRLF
	    cScript += "PROGRAMA        NVARCHAR(255) NOT NULL,"+CRLF
	    cScript += "DATA_LOG        DATETIMEOFFSET DEFAULT SYSDATETIMEOFFSET() NOT NULL,"+CRLF
	    cScript += "ORIGEM          NVARCHAR(255) NOT NULL,"+CRLF
        cScript += "ENDPOINT        NVARCHAR(2048) NOT NULL,"+CRLF
	    cScript += "METODO          NVARCHAR(10) NOT NULL,"+CRLF
	    cScript += "REQUEST_BODY    NVARCHAR(MAX),"+CRLF
	    cScript += "RESPONSE_BODY   NVARCHAR(MAX),"+CRLF
	    cScript += "STATUS_CODE     SMALLINT,"+CRLF
	    cScript += "TEMPO_RESPOSTA  INT, "+CRLF
	    cScript += "USUARIO         NVARCHAR(255),"+CRLF
	    cScript += "AMBIENTE        NVARCHAR(50) NOT NULL,"+CRLF
	    cScript += "OBSERVACAO      NVARCHAR(MAX)"+CRLF
	    cScript += ");"+CRLF
	    If TCSqlExec(cScript) < 0
	    	conout("Não foi possivel criar tabela de log, Motivo: "+TCSQLError() )
	    endif
	    cScript:= "CREATE INDEX idx_apilog_data ON API_LOG (DATA_LOG DESC);"
	    If TCSqlExec(cScript) < 0
	    	conout("Não foi possivel criar INDICE "+cScript+", Motivo: "+TCSQLError() )
	    endif
	    cScript:= "CREATE INDEX idx_apilog_endpoint ON API_LOG (ENDPOINT);"
	    If TCSqlExec(cScript) < 0
	    	conout("Não foi possivel criar INDICE "+cScript+", Motivo:  "+TCSQLError() )
	    endif
	    cScript:= "CREATE INDEX idx_apilog_status_code ON API_LOG (STATUS_CODE);"
	    If TCSqlExec(cScript) < 0
	    	conout("Não foi possivel criar INDICE "+cScript+", Motivo: "+TCSQLError() )
	    endif
	    cScript:= "CREATE INDEX idx_apilog_id_endpoint ON API_LOG (ID_LOG, ENDPOINT);"
	    If TCSqlExec(cScript) < 0
	    	conout("Não foi possivel criar INDICE "+cScript+", Motivo: "+TCSQLError() )
	    endif
	    cScript:= "CREATE INDEX idx_apilog_id_data ON API_LOG (ID_LOG, DATA_LOG);"
	    If TCSqlExec(cScript) < 0
	    	conout("Não foi possivel criar INDICE "+cScript+", Motivo: "+TCSQLError() )
	    endif
	    cScript:= "CREATE INDEX idx_apilog_id_status ON API_LOG (ID_LOG, STATUS_CODE);"
	    If TCSqlExec(cScript) < 0
	    	conout("Não foi possivel criar INDICE "+cScript+", Motivo: "+TCSQLError() )
	    endif
    endif

return
