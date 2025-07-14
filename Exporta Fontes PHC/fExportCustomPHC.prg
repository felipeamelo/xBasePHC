**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fExportCustomPHC()
**//|Autor.....: Felipe Aurélio de Melo
**//|Data......: Julho de 2025
**//|Descricao.: Exporta customizações para ficheiros .prg
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
Local i, aTabsVFP, cQueryTmp, cQryLin12, cQryLin13, cQryLin14
Dimension aTabsVFP[14,8]

* Cria cursor que será utilizado para exportar os ficheiros 
fecha("dbResultTmp")
CREATE CURSOR dbResultTmp (;
        tabela_nome C( 20), ;
        tabela_desc C( 30), ;
        inactivo    C(  1), ;
        id          C(  5), ;
        codigo_vfp  M     , ;
        resumo      C(250), ;
        descricao   C(100), ;
        query       C(250) ;
)

**Query executado na linha 12 abaixo
**Como query é mais complexa, então criei ela aqui
TEXT TO cQryLin12 NOSHOW TEXTMERGE
    SELECT 
        DENSE_RANK() OVER (ORDER BY formC.ecran, formC.dfdidoc) AS id,
        IIF((select count(*) from formB where formB.inactivo = 1 and formB.ecran = formC.ecran and formB.dfdidoc = formC.dfdidoc)>0,1,IIF((select count(*) from pdu where pdu.inactivo = 1 and pdu.sysnmpdu = formC.ecran)>0,1,0)) as inactivo_new,
        formC.ecran,
        formC.dfdidoc,
        formC.objpai,
        formC.numpag, 
        formC.ordem,
        formC.nome,
        formC.prop
    FROM 
        formC(nolock)
    ORDER BY formC.ecran, formC.dfdidoc, formC.objpai, formC.numpag, formC.ordem
ENDTEXT


**Query executado na linha 13 abaixo
**Como query é mais complexa, então criei ela aqui
TEXT TO cQryLin13 NOSHOW TEXTMERGE
    SELECT 
        0 as inactivo_new,
        ROW_NUMBER() OVER (ORDER BY o.name) as id,
        o.name AS ProcedureName,
        cast(m.definition as TEXT) as codigo_tsql
    FROM sys.sql_modules(nolock) m
    JOIN sys.objects(nolock) o ON m.object_id = o.object_id
    WHERE o.type = 'P'  -- 'P' = Stored Procedure
    AND m.definition > ''
    order by o.name
ENDTEXT


**Query executado na linha 13 abaixo
**Como query é mais complexa, então criei ela aqui
TEXT TO cQryLin14 NOSHOW TEXTMERGE
    SELECT 
        0 as inactivo_new,
        ROW_NUMBER() OVER (ORDER BY o.name) as id,
        o.name AS ViewName,
        cast(m.definition as TEXT) as codigo_tsql
    FROM sys.sql_modules(nolock) m
    JOIN sys.objects(nolock) o ON m.object_id = o.object_id
    WHERE o.type = 'V'  -- 'V' = View
    AND m.definition > ''
    order by o.name
ENDTEXT


* Preenche o array com os detalhes de cada tabela que será utilizada para exportar o codigo fonte
*//               , (1)Tabela  , (2)Descrição Tabela          , (3)Inactivo     , (4)ID          , (5)Código Fonte                                                               , (6)resumo                                                                                                                                             , (7)descricao                                                                            , (8)Query 
fAdd(01 ,@aTabsVFP, "TC+TI"    , "Teclas de Utilizador"       , "inactivo_new"  , "tcid"         , "dbCursorTmp.retorno"                                                         , "upper(alltrim(dbCursorTmp.ecran))      +' ( '+alltrim(dbCursorTmp.teclas)    +' ) - '+alltrim(dbCursorTmp.descricao)"                                , "alltrim(dbCursorTmp.descricao)"                                                        ,"select CASE WHEN tc.inactivo = 1 THEN 1 ELSE CAST(ti.inactivo AS INT) END AS inactivo_new,tcid,retorno,ecran,teclas,descricao, DENSE_RANK() OVER (ORDER BY ecran, teclas) as id from tc(nolock) inner join ti(nolock) on ti.tcstamp = tc.tcstamp" )
fAdd(02 ,@aTabsVFP, "EVENTOS"  , "Eventos de Utilizador"      , "inactivo_new"  , "eventosid"    , "fConcat('x','expressao','condicao','emailsu','emailte','emailto','emailto')" , "upper(alltrim(dbCursorTmp.ecran))      +' ( '+alltrim(dbCursorTmp.qevento)   +' ) - '+alltrim(dbCursorTmp.resumo)"                                   , "alltrim(dbCursorTmp.resumo)"                                                           ,"select CAST(inactivo AS INT) AS inactivo_new, eventosid, ecran, qevento, resumo, expressao, condicao, emailsu, emailte, emailto, emailto from eventos(nolock)" )
fAdd(03 ,@aTabsVFP, "INTU"     , "Instruções Internas"        , "inactivo_new"  , "intuid"       , "dbCursorTmp.xbase"                                                           , "upper(alltrim(dbCursorTmp.ecran))      +' ( '+alltrim(dbCursorTmp.nome)      +' ) '"                                                                 , "alltrim(dbCursorTmp.ajuda)"                                                            ,"select CAST(inactivo AS INT) AS inactivo_new, intuid, xbase, nome, ajuda, ecran from intu(nolock)" )
fAdd(04 ,@aTabsVFP, "BRULE"    , "Regras de Utilizador"       , "inactivo_new"  , "bruleid"      , "dbCursorTmp.expressao"                                                       , "upper(alltrim(dbCursorTmp.tabela))     +' ( '+alltrim(dbCursorTmp.descricao) +' ) '"                                                                 , "alltrim(dbCursorTmp.mensagem)"                                                         ,"select CAST(inactivo AS INT) AS inactivo_new, bruleid, tabela, expressao, descricao, mensagem, DENSE_RANK() OVER (ORDER BY bruleid, descricao) id from brule(nolock) order by bruleid" )
fAdd(05 ,@aTabsVFP, "UDEFS"    , "Valores por Defeito"        , "inactivo_new"  , "udefsid"      , "fConcat('x','xbase','transact')"                                             , "upper(alltrim(dbCursorTmp.campo))      +' ( '+alltrim(dbCursorTmp.descricao) +' ) '"                                                                 , "alltrim(dbCursorTmp.descricao)"                                                        ,"select CAST(inactivo AS INT) AS inactivo_new, udefsid, campo, descricao, xbase, transact from udefs(nolock) order by campo" )
fAdd(06 ,@aTabsVFP, "USQL"     , "Análises Avanç Utilizador"  , "inactivo_new"  , "usqlid"       , "fConcat('x','sqlexpr','ssxbase')"                                            , "alltrim(dbCursorTmp.descricao)         +' ( '+alltrim(dbCursorTmp.grupo)     +' ) '"                                                                 , "alltrim(dbCursorTmp.descricao)"                                                        ,"select CAST(inactivo AS INT) AS inactivo_new, usqlid, descricao, sqlexpr, grupo, ssxbase, DENSE_RANK() OVER (ORDER BY usqlid, descricao) as id from usql(nolock) order by usqlid" )
fAdd(07 ,@aTabsVFP, "FTU"      , "Funções TSQL de Utilizador" , "inactivo_new"  , "id"           , "dbCursorTmp.xbase"                                                           , "alltrim(dbCursorTmp.nome)"                                                                                                                           , "alltrim(dbCursorTmp.nome)"                                                             ,"select CAST(inactivo AS INT) AS inactivo_new, ROW_NUMBER() OVER (ORDER BY nome) as id, nome, xbase from FTU(nolock) " )
fAdd(08 ,@aTabsVFP, "FXBASE"   , "Funções de Utilizador"      , "inactivo_new"  , "id"           , "dbCursorTmp.xbase"                                                           , "alltrim(dbCursorTmp.nome)              +' ( '+alltrim(dbCursorTmp.titulo)    +' ) '"                                                                 , "alltrim(dbCursorTmp.titulo)"                                                           ,"select CAST(inactivo AS INT) AS inactivo_new, ROW_NUMBER() OVER (ORDER BY nome) as id, nome, titulo, xbase from FXBASE(nolock) " )
fAdd(09 ,@aTabsVFP, "AL"       , "Alertas do Utilizador"      , "inactivo_new"  , "alid"         , "fConcat('x','condicao','emailsu','emailte','emailto','expressao')"           , "alltrim(dbCursorTmp.resumo)"                                                                                                                         , "alltrim(dbCursorTmp.resumo)"                                                           ,"select CAST(inactivo AS INT) AS inactivo_new, alid, resumo, condicao, emailsu, emailte, emailto, expressao from al(nolock)" )
fAdd(10 ,@aTabsVFP, "USTR"     , "Triggers de Utilizador"     , "inactivo_new"  , "ustrid"       , "dbCursorTmp.tsql"                                                            , "upper(alltrim(dbCursorTmp.tabela))     +' ( '+alltrim(dbCursorTmp.titulo)    +' ) '"                                                                 , "alltrim(dbCursorTmp.titulo)"                                                           ,"select CASE WHEN activo = 1 THEN 0 ELSE 1 END AS inactivo_new, ustrid, tabela, titulo, tsql from ustr(nolock)" )
fAdd(11 ,@aTabsVFP, "OUN"      , "Opções de Navegação"        , "inactivo_new"  , "ounid"        , "dbCursorTmp.xbase"                                                           , "upper(alltrim(dbCursorTmp.grupo))      +' ( '+alltrim(dbCursorTmp.nome)      +' ) '"                                                                 , "alltrim(dbCursorTmp.nome)"                                                             ,"select CAST(inactivo AS INT) AS inactivo_new, ounid, nome, grupo, ecran, xbase, nmother from OUN(nolock) where len(rtrim(CAST(xbase AS VARCHAR))) > 5" )
fAdd(12 ,@aTabsVFP, "FORMC"    , "Ecrans e Paineis Inform"    , "inactivo_new"  , "id"           , "dbCursorTmp.prop"                                                            , "upper(alltrim(dbCursorTmp.ecran))      +' ( '+aStr(dbCursorTmp.dfdidoc)      +' ) - '+alltrim(dbCursorTmp.objpai) + ':'+alltrim(dbCursorTmp.nome)"   , "Upper(alltrim(dbCursorTmp.ecran)) + ' - (idDoc.' +aStr(dbCursorTmp.dfdidoc) + ')' "    ,"cQryLin12" )
fAdd(13 ,@aTabsVFP, "SQL_SP"   , "SQLServer Procedure"        , "inactivo_new"  , "id"           , "dbCursorTmp.codigo_tsql"                                                     , "alltrim(dbCursorTmp.ProcedureName)"                                                                                                                  , "alltrim(dbCursorTmp.ProcedureName)"                                                    ,"cQryLin13" )
fAdd(14 ,@aTabsVFP, "SQL_VIEW" , "SQLServer View"             , "inactivo_new"  , "id"           , "dbCursorTmp.codigo_tsql"                                                     , "alltrim(dbCursorTmp.ViewName)"                                                                                                                       , "alltrim(dbCursorTmp.ViewName)"                                                         ,"cQryLin14" )


&& Tabela		Titulo     
**********************************************************************
* x  - Campos        Campos do Utilizador      
* x  - DIC           Dicionário                                                                                                                                      	
* ok - TC/TI         Teclas de Utilizador                                                                                                                                  	
* ok - EVENTOS       Eventos de Utilizador                                                                                                                                 	
* ok - INTU          Instruções Internas                                                                                                                                   	
* ok - BRULE         Regras de Utilizador                                                                                                                                  	
* ok - UDEFS         Valores por Defeito                                                                                                                                   	
* ok - USQL          Análises Avançadas do Utilizador                                                                                                                      	
* ok - FTU           Funções TSQL de Utilizador    
* ok - FXBASE        Funções de Utilizador   
* ok - AL            Alertas do Utilizador
* ok - USTR          Triggers ou Indices de Utilizador      
* ok - OUN           Opções de Navegação
*    - PEM           Propriedade de Ecrã  
*    - PDU           Painel de Informação 
*    - XFT           Configuração de Faturação Personalizada	
*    - NTF           Notificação de Utilizador  
*    - ULE/ULEL      Listagens Rápida de Utilizador 
* ------------------------------------------------------------------- 
* ATE AQUI ----------------------------------------------------------
* ------------------------------------------------------------------- 
* FORMC         Ecrans
* formb         Ecrans Personalizados
* tufo          Ecrans do Utilizador
* js            Ligações entre ecrans
* tu            Tabelas de utilizador
* tuca          Tabelas de utilizador (campos)
* lt            Ligações entre tabelas
* feu           Filtros de Utilizador
* parau         Parametros do utilizador
* FTIDUC        IDU das Faturas
* FTIDUL        IDU das Faturas
* BOIDUC        IDU das Dossier
* BOIDUL        IDU das Dossier
* FLTFT         filtros personalizados das Faturas
* FLTBO         filtros personalizados dos Dossier
**********************************************************************


* Percorre array das tabelas acima para gerar os ficheiros com o codigo fonte
For i = 1 To aLen(aTabsVFP, 1)
    cQueryTmp = ""
    cQueryTmp = IIf(SubStr(aTabsVFP[i,8],1,7)=="cQryLin",Evaluate(aTabsVFP[i,8]),aTabsVFP[i,8])
    fecha("dbCursorTmp")
    u_sqlexec(cQueryTmp,"dbCursorTmp")
    If Used("dbCursorTmp") And Reccount("dbCursorTmp") > 0
        SELECT dbCursorTmp
        SCAN
            SELECT dbResultTmp
            GO BOTTOM
            APPEND BLANK
            REPLACE dbResultTmp.tabela_nome  WITH aTabsVFP[i,1]
            REPLACE dbResultTmp.tabela_desc  WITH aTabsVFP[i,2]
            REPLACE dbResultTmp.inactivo     WITH aStr(Evaluate("dbCursorTmp."+aTabsVFP[i,3]))
            REPLACE dbResultTmp.id           WITH aStr(Evaluate("dbCursorTmp."+aTabsVFP[i,4]))
            REPLACE dbResultTmp.codigo_vfp   WITH      Evaluate(aTabsVFP[i,5])
            REPLACE dbResultTmp.resumo       WITH      Evaluate(aTabsVFP[i,6])
            REPLACE dbResultTmp.descricao    WITH      Evaluate(aTabsVFP[i,7])
            REPLACE dbResultTmp.query        WITH aTabsVFP[i,8]
        ENDSCAN
    EndIf
    fecha("dbCursorTmp")

EndFor

* Mostra resultado em tela
mostrameisto("dbResultTmp")

*Exporta customizações para ficheiros .prg
**fExportPRG()
**Return

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fExportPRG()
**//|Autor.....:
**//|Data......: 
**//|Descricao.: Exporta customizações para ficheiros .prg
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
**Function fExportPRG()
**-----------------------------------------------------------*
LOCAL oDict, lcKey, nQtde, lcConteudoAcumulado, lcFilename
oDict = CREATEOBJECT("Scripting.Dictionary")
nQtde = 0

SELECT dbResultTmp
SCAN
    lcKey = Alltrim(dbResultTmp.tabela_nome)+'-'+Alltrim(dbResultTmp.id)+'-'+Alltrim(IIf(dbResultTmp.tabela_nome="FORMC",dbResultTmp.descricao,dbResultTmp.resumo))
    IF oDict.Exists(lcKey)
        oDict.Item(lcKey) = oDict.Item(lcKey) + CHR(13)+CHR(10) + fConcat("ok",dbResultTmp.codigo_vfp, dbResultTmp.resumo)
    ELSE
        oDict.Add(lcKey, lcKey + CHR(13)+CHR(10) + fConcat("ok",dbResultTmp.codigo_vfp, dbResultTmp.resumo))
    ENDIF
ENDSCAN

* Exportar por ID
FOR EACH lcKey IN oDict.Keys
    SELECT dbResultTmp
    LOCATE FOR ALLTRIM(tabela_nome)+'-'+ALLTRIM(id)+'-'+ALLTRIM(IIf(tabela_nome="FORMC",descricao,resumo)) == lcKey

    * Garante pasta
    IF NOT DIRECTORY("C:\resultados\")
        MD "C:\resultados\"
    ENDIF

    * Construção do nome do ficheiro
    lcFilename = IIF(dbResultTmp.inactivo=="1","!","") + ;
                 ALLTRIM(dbResultTmp.tabela_desc) + ;
                 " [id." + Alltrim(dbResultTmp.id) + "] - " + ;
                 fTrataTexto(IIf(dbResultTmp.tabela_nome="FORMC",dbResultTmp.descricao,dbResultTmp.resumo)) + ".prg"

    lcFilename = FORCEPATH(lcFilename, "C:\resultados\")

    * Salvar conteúdo acumulado
    lcConteudoAcumulado = oDict.Item(lcKey)
    STRTOFILE(lcConteudoAcumulado, lcFilename)
ENDFOR

Msg("Os ficheiros PRG foram criados em: "+Chr(13)+Chr(10)+"C:\resultados\")

Return
**EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fAdd()
**//|Autor.....:
**//|Data......: 
**//|Descricao.: Define a função fAdd para adicionar uma linha ao array
**//|Observacao:
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fAdd(nRow, arrTarget, v1, v2, v3, v4, v5, v6, v7, v8)
**-----------------------------------------------------------*
    * A função adiciona os valores passados para a linha nRow
    arrTarget[nRow, 1] = v1
    arrTarget[nRow, 2] = v2
    arrTarget[nRow, 3] = v3
    arrTarget[nRow, 4] = v4
    arrTarget[nRow, 5] = v5
    arrTarget[nRow, 6] = v6
    arrTarget[nRow, 7] = v7
    arrTarget[nRow, 8] = v8
EndFunc


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fConcat()
**//|Autor.....:
**//|Data......: 
**//|Descricao.: Trata concatenação
**//|Observacao:
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fConcat(cTipo,cVarTmp, v1, v2, v3, v4, v5, v6, v7, v8)
**-----------------------------------------------------------*
Local cRet, cQbra
cRet = ""
cQbra = CHR(13)+CHR(10)

If cTipo == 'ok'
    **//Trata para remover aspas desnecessária
    cVarTmp = StrTran(cVarTmp,cQbra+CHR(34)           , "")
    cVarTmp = StrTran(cVarTmp,CHR(10)+CHR(34)         , "")
    cVarTmp = StrTran(cVarTmp,CHR(13)+CHR(34)         , "")
    cVarTmp = StrTran(cVarTmp,CHR(34)+cQbra+CHR(60)   , cQbra)
    cVarTmp = StrTran(cVarTmp,':"<<'                  , ':<<')

    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + fTrataTexto(v1) + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + cVarTmp
    Return(cRet)
EndIf

If !Empty(v1)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v1 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v1)
EndIf

If !Empty(v2)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v2 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v2)
EndIf

If !Empty(v3)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v3 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v3)
EndIf

If !Empty(v4)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v4 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v4)
EndIf

If !Empty(v5)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v5 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v5)
EndIf

If !Empty(v6)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v6 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v6)
EndIf

If !Empty(v7)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v7 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v7)
EndIf

If !Empty(v8)
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + "#### campo: " + v8 + " ####"
    cRet = cRet + cQbra + "======================================================="
    cRet = cRet + cQbra + Evaluate("dbCursorTmp."+v8)
EndIf

Return(cRet)
EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fTrataTexto()
**//|Autor.....:
**//|Data......: 
**//|Descricao.: Trata texto removendo caractere especial
**//|Observacao:
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fTrataTexto(tcTexto)
**-----------------------------------------------------------*
LOCAL i, lcChar, lcResultado, lnCod, lcSemAcentos
lcResultado = ""

* Remove espaços
lcSemAcentos = Alltrim(tcTexto)

* Etapa 1: Substitui acentos por letras simples
lcSemAcentos = STRTRAN(lcSemAcentos, "á", "a")
lcSemAcentos = STRTRAN(lcSemAcentos, "à", "a")
lcSemAcentos = STRTRAN(lcSemAcentos, "â", "a")
lcSemAcentos = STRTRAN(lcSemAcentos, "ã", "a")
lcSemAcentos = STRTRAN(lcSemAcentos, "ä", "a")
lcSemAcentos = STRTRAN(lcSemAcentos, "é", "e")
lcSemAcentos = STRTRAN(lcSemAcentos, "è", "e")
lcSemAcentos = STRTRAN(lcSemAcentos, "ê", "e")
lcSemAcentos = STRTRAN(lcSemAcentos, "í", "i")
lcSemAcentos = STRTRAN(lcSemAcentos, "ó", "o")
lcSemAcentos = STRTRAN(lcSemAcentos, "ô", "o")
lcSemAcentos = STRTRAN(lcSemAcentos, "õ", "o")
lcSemAcentos = STRTRAN(lcSemAcentos, "ú", "u")
lcSemAcentos = STRTRAN(lcSemAcentos, "ç", "c")

* Também tratar letras maiúsculas
lcSemAcentos = STRTRAN(lcSemAcentos, "Á", "A")
lcSemAcentos = STRTRAN(lcSemAcentos, "À", "A")
lcSemAcentos = STRTRAN(lcSemAcentos, "Â", "A")
lcSemAcentos = STRTRAN(lcSemAcentos, "Ã", "A")
lcSemAcentos = STRTRAN(lcSemAcentos, "É", "E")
lcSemAcentos = STRTRAN(lcSemAcentos, "Ê", "E")
lcSemAcentos = STRTRAN(lcSemAcentos, "Í", "I")
lcSemAcentos = STRTRAN(lcSemAcentos, "Ó", "O")
lcSemAcentos = STRTRAN(lcSemAcentos, "Ô", "O")
lcSemAcentos = STRTRAN(lcSemAcentos, "Õ", "O")
lcSemAcentos = STRTRAN(lcSemAcentos, "Ú", "U")
lcSemAcentos = STRTRAN(lcSemAcentos, "Ç", "C")

* Também tratar troca de caracter especial
lcSemAcentos = STRTRAN(lcSemAcentos, ":", ".")

* Etapa 2: Remove caracteres inválidos
FOR i = 1 TO LEN(lcSemAcentos)
    lcChar = SUBSTR(lcSemAcentos, i, 1)
    lnCod = ASC(lcChar)

    * Letras, números, hífen, underscore
    IF (lnCod >= 65 AND lnCod <= 90) OR ;
        (lnCod >= 97 AND lnCod <= 122) OR ;
        (lnCod >= 48 AND lnCod <= 57) OR ;
        lnCod = 45 OR lnCod = 95 OR lnCod = 32 OR lnCod = 40 OR lnCod = 41 OR lnCod = 46 && - ou _ ou espaço ou ( ou ) ou .
        lcResultado = lcResultado + lcChar
    ENDIF
ENDFOR

RETURN Alltrim(lcResultado)

ENDFUNC


