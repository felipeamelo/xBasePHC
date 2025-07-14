* Define o array unidimensional para armazenar os servidores como strings
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fBaixaPODdoFTP()
**//|Autor.....:
**//|Data......: 14 de novembro de 2024, 09:00
**//|Descricao.:
**//|Observacao: Baixa PODs do FTP dos Agentes
**//+-----------------------------------------------------------------------------------//
Local i, oFtp, aServidores, cFolderDest, cResumoLOG, cMensagemLOG
 
Dimension aServidores[17,7]
 
cFolderDest = "\\192.168.1.174\servicos\EDI\TrackAndTrace\DownPODTest\"
 
* Preenche o array com os detalhes de cada servidor como strings
fAdd(01 ,@aServidores, "BBL"       , "bblpod.solulog.fr"                 , 21,   "luso-porto"           , "djJCQ25DRDhFU3NkYy1ET1M4b0g=" , "/out/"                               )
fAdd(02, @aServidores, "TIF Lyon"  , "bblpod.solulog.fr"                 , 21,   "luso-lisbonne-26"     , "bEhiandJUjYyRGhOQWFCbXhDWm0=" , "/out/"                               )
fAdd(03, @aServidores, "Lagermax"  , "dw.lagermax.com"                   , 21,   "lusocargo"            , "cUgzR2lTdEY="                 , "/podout/"                            )
fAdd(04, @aServidores, "Lombard"   , "157.231.112.126"                   , 22,   "Lusocargo-LombardFTP" , "Slh4NSUzSW9VMw=="             , "/Lusocargo/OUT/POD/"                 )
fAdd(05, @aServidores, "TISA"      , "212.203.89.11"                     , 9022, "lusocargo"            , "TMK1c0BUIXNhXzIwMjM="         , "/out/"                               )
fAdd(06, @aServidores, "Raben"     , "sftp2.raben-group.com"             , 22,   "lusocargo"            , "ejVWZnNHMXNEVGVhRjVnNA=="     , "/PROD/DPT45/PODout/"                 )
fAdd(07, @aServidores, "Seitrans"  , "ftp.seitrans.es"                   , 21,   "Mlusocargo"           , "U2VpTWFkMjAxMQ=="             , "/Pod_Out/"                           )
fAdd(08, @aServidores, "UVB"       , "ftpedi.ubv.it"                     , 21,   "lusoftp"              , "d1JlelV4MkRyaQ=="             , "/out/Porto_BO/POD/"                  )
fAdd(09, @aServidores, "UVB"       , "ftpedi.ubv.it"                     , 21,   "lusoftp"              , "d1JlelV4MkRyaQ=="             , "/out/Porto_CO/POD/"                  )
fAdd(10, @aServidores, "ITALTRA"   , "ftpedi.italtransport.it"           , 22,   "ital-lusocargo"       , "MUM2NnVyZyFkc1Q="             , "/PORTO/IN/POD/"                      )
fAdd(11, @aServidores, "Woehl"     , "edi.heppner.fr"                    , 21,   "lusocargo"            , "dEI4eE9JTG1hSUVB"             , "/POD/out/"                           )
fAdd(12, @aServidores, "Ziegler"   , "5.9.153.194"                       , 21,   "ziegler"              , "S3BlNjc/ZTUy"                 , "/to_lusocargo/"                      )
fAdd(13, @aServidores, "Delamode"  , "5.9.153.194"                       , 21,   "delamode"             , "WGw0NEB1ejI"                  , "/to_lusocargo/POD/"                  )
fAdd(14, @aServidores, "Cargoline" , "ftp.cepra.de"                      , 22,   "9082"                 , "UHVzSkVPYlNPcTlreUE="         , "/DOCUMENTS/OUT/POD/"                 )
fAdd(15, @aServidores, "Blue Water", "ftp.bws.dk"                        , 22,   "lusocargo"            , "MnhkQjJiMG9ScjYkVXJVNk5JeE8=" , "/tolusocargo/pod/"                   )
fAdd(16, @aServidores, "Glob Sped" , "edi-prod.schneider-transport.com"  , 22,   "sftp_LUSO_PROD"       , "a2o1WjgyVlFRZQ=="             , "/home/sftp_LUSO_PROD/out/Porto/POD/" )
fAdd(17, @aServidores, "Van Duuren", "sftp.vanduuren.com"                , 22,   "Lusocargoporto"       , "QHIoOjRQMnM="                 , "/out/POD/"                           )
 
* Itera sobre cada servidor para acessar, decodificar a senha, e processar arquivos
For i = 1 To ALEN(aServidores, 1)  && Usa ALEN para obter o número de linhas no array
 
    * Decodifica a senha Base64 (quinta coluna)
    lcPassFtp = STRCONV(ALLTRIM(aServidores[i, 5]), 14)
 
    * Conectar ao FTP ou SFTP com base na porta
    If fConectaFTP_SFTP(aServidores[i, 1], aServidores[i, 2], aServidores[i, 3], aServidores[i, 4], lcPassFtp, @oFtp)
       
        * Confirmação de que a conexão foi bem-sucedida
        * *MESSAGEBOX("Conexão estabelecida com sucesso. Preparando para baixar arquivos em " + aServidores[i, 6])
 
        If aServidores[i, 3] == 21
            fDwn21POD(@oFtp, aServidores[i, 1], aServidores[i, 6], cFolderDest)
        EndIf
 
        If aServidores[i, 3] == 22 or aServidores[i, 3] == 9022
            fDwn22POD(@oFtp, aServidores[i, 1], aServidores[i, 6], cFolderDest)
        EndIf
    Else
        * *MESSAGEBOX("Erro ao conectar no servidor " + aServidores[i, 1])
    EndIf
EndFor
 
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fConectaFTP_SFTP()
**//|Autor.....:
**//|Data......: 14 de novembro de 2024, 09:00
**//|Descricao.:
**//|Observacao:
**//+-----------------------------------------------------------------------------------//
Function fConectaFTP_SFTP(cClientFTP, lcHostname, lnPort, lcUserFtp, lcPassFtp, loFtp)
    Local lOk
 
    * Verificar porta para escolher FTP ou SFTP
    If lnPort == 21
        * Conectar via FTP
        loFtp = CreateObject('Chilkat_9_5_0.Ftp2')
       
        * Configurar timeout de conexão e leitura
        loFtp.ConnectTimeout = 60    && 60 segundos
        loFtp.ReadTimeout = 60000    && 60 segundos para operações de leitura
 
        * *MESSAGEBOX("Tentativa de criar objeto FTP para " + cClientFTP + ". Tipo de objeto: " + VARTYPE(loFtp))
 
        If VARTYPE(loFtp) <> "O"
            * *MESSAGEBOX("Falha ao criar objeto FTP para " + cClientFTP)
            Return .F.
        EndIf
 
        loFtp.Hostname = lcHostname
        loFtp.Username = lcUserFtp
        loFtp.Password = lcPassFtp
        loFtp.Port = lnPort
        loFtp.Passive = 1
        lOk = IIF(loFtp.Connect() = 1, .T., .F.)
       
        If lOk
            * *MESSAGEBOX("Conexão FTP bem-sucedida para " + cClientFTP + " no host " + lcHostname)
        Else
            * *MESSAGEBOX("Falha na conexão FTP para " + cClientFTP + " no host " + lcHostname + ". Erro: " + loFtp.LastErrorText)
        EndIf
 
    Else
        If lnPort == 22 Or lnPort == 9022
            * Conectar via SFTP
            loFtp = CreateObject('Chilkat_9_5_0.SFtp')
            loFtp.ConnectTimeoutMs = 60000  && 60 segundos para SFTP
 
            * *MESSAGEBOX("Tentativa de criar objeto SFTP para " + cClientFTP + ". Tipo de objeto: " + VARTYPE(loFtp))
 
            If VARTYPE(loFtp) <> "O"
                * *MESSAGEBOX("Falha ao criar objeto SFTP para " + cClientFTP)
                Return .F.
            EndIf
 
            lOk = IIF(loFtp.Connect(lcHostname, lnPort) = 1, .T., .F.)
 
            * Configura o usuário e a senha após conectar
            If lOk
                loFtp.AuthenticatePw(lcUserFtp, lcPassFtp)
                If loFtp.InitializeSftp() <> 1
                    * *MESSAGEBOX("Falha na inicialização do SFTP para " + cClientFTP + " no host " + lcHostname + ". Erro: " + loFtp.LastErrorText)
                    lOk = .F.
                Else
                    * *MESSAGEBOX("Conexão SFTP bem-sucedida para " + cClientFTP + " no host " + lcHostname)
                EndIf
            Else
                * *MESSAGEBOX("Falha na conexão SFTP para " + cClientFTP + " no host " + lcHostname + ". Erro: " + loFtp.LastErrorText)
            EndIf
        Else
            * Porta inválida ou não suportada
            * *MESSAGEBOX("Porta " + TRANSFORM(lnPort) + " não suportada para " + cClientFTP)
            lOk = .F.
        EndIf
    EndIf
 
    Return lOk
EndFunc
 
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fDwn21POD()
**//|Autor.....:
**//|Data......: 14 de novembro de 2024, 09:00
**//|Descricao.:
**//|Observacao:
**//+-----------------------------------------------------------------------------------//
Function fDwn21POD(loFtp, cClientFTP, lcDirRemote, lcDirLocal)
    Local i, lnSuccess, lcRemoteFileName, nFiles, nPdfFiles, lOk, nTamNew
    lOk = .T.
 
    * *MESSAGEBOX("Entrou na função fDwn21POD para " + cClientFTP + ". Tipo de objeto loFtp: " + VARTYPE(loFtp))
 
    If VARTYPE(loFtp) = "O"
        * *MESSAGEBOX("O objeto loFtp é do tipo objeto (O)")
       
        nFiles = 0
 
        **//Tentativa 1
        If nFiles == 0
            * Configura o diretório remoto sem um padrão para capturar todos os arquivos
            loFtp.ChangeRemoteDir(lcDirRemote)
            loFtp.ListPattern = "./*.*"  && Lista todos os arquivos e diretórios
            nFiles = loFtp.GetDirCount()
        EndIf
 
        **//Tentativa 2
        If nFiles == 0
            * Configura o diretório remoto sem um padrão para capturar todos os arquivos
            loFtp.ChangeRemoteDir(lcDirRemote)
            loFtp.ListPattern = "*.*"  && Lista todos os arquivos e diretórios
            nFiles = loFtp.GetDirCount()
        EndIf
 
 
        nPdfFiles = 0  && Contador de arquivos com extensão .pdf
        If nFiles > 0
            * Contagem de arquivos PDF, independente de maiúsculas ou minúsculas
            For i = 0 To nFiles - 1
                If loFtp.GetIsDirectory(i) <> 1
                    lcRemoteFileName = AllTrim(loFtp.GetFilename(i))
 
                    * Verifica se o arquivo termina com .pdf (em qualquer caso de letras)
                    If LOWER(RIGHT(lcRemoteFileName, 4)) $ ".pdf/.jpg/.jpeg/.png/.tif/"
                        nPdfFiles = nPdfFiles + 1
                    EndIf
                EndIf
            EndFor
 
            * Exibe a contagem de arquivos PDF encontrados
            If nPdfFiles > 0
                *MESSAGEBOX(cClientFTP+": Número de arquivos PDF/PNG/JPG encontrados: " + TRANSFORM(nPdfFiles))
            Else
                *MESSAGEBOX("Nenhum arquivo PDF encontrado em " + lcDirRemote + " no servidor " + cClientFTP)
                RELEASE loFtp
                Return .F.
            EndIf
        Else
            *MESSAGEBOX("Nenhum arquivo encontrado em " + lcDirRemote + " no servidor " + cClientFTP + ". Erro: " + loFtp.LastErrorText)
            RELEASE loFtp
            Return .F.
        EndIf
 
        * Verifica se o diretório local existe
        If !DIRECTORY(lcDirLocal)
            *MESSAGEBOX("Erro: O diretório local " + lcDirLocal + " não existe ou não é acessível.")
            RELEASE loFtp
            Return .F.
        EndIf
 
        * Tenta baixar cada arquivo PDF encontrado
        For i = 0 To nFiles - 1
            If loFtp.GetIsDirectory(i) <> 1
                lcRemoteFileName = AllTrim(loFtp.GetFilename(i))
                loDt = loFtp.GetLastModDt(i)
                lcRemoteFileDateTime = StrTran(SubStr(loDt.GetAsTimestamp(0),1,10),"-","")
                nTamNew = loFtp.GetSizeByName(i)
 
                * Verifica se o arquivo termina com ".pdf" de forma insensível a maiúsculas e minúsculas
                If LOWER(RIGHT(lcRemoteFileName, 4)) $ ".pdf/.jpg/.jpeg/.png/.tif/" And lcRemoteFileDateTime >= DtoS(date() -5)
                    If fNewFile(cClientFTP, lcDirRemote, lcRemoteFileName)
                        * Caminho completo de download local
                        Local lcLocalPath
                        lcLocalPath = lcDirLocal + lcRemoteFileName
   
                        * Tenta o download do arquivo
                        lnSuccess = loFtp.GetFile(lcDirRemote + lcRemoteFileName, lcLocalPath)
   
                        If lnSuccess = 1
                            fLogFileOk("ALERTA128", lcDirLocal, lcRemoteFileName)
                            * *MESSAGEBOX("Sucesso no download de " + lcRemoteFileName + " para " + lcLocalPath)
                        Else
                            cResumoLOG = "Erro ao realizar query from u_impstat512"
                            cMensagemLOG = "Falha no download de " + lcRemoteFileName + ". Erro: " + loFtp.LastErrorText
                            fGravaLog( 128, "processarArquivos TFT " ,cResumoLOG ,cMensagemLOG )
                            lOk = .F.
                        EndIf
                    EndIf
                EndIf
                RELEASE loDt
            EndIf
        EndFor
    Else
        *MESSAGEBOX("loFtp não é um objeto válido em fDwn21POD")
        RELEASE loFtp
        Return .F.
    EndIf
 
    RELEASE loFtp
    Return lOk
EndFunc
 
 
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fDwn22POD()
**//|Autor.....:
**//|Data......: 14 de novembro de 2024, 09:00
**//|Descricao.:
**//|Observacao:
**//+-----------------------------------------------------------------------------------//
Function fDwn22POD(loSftp, cClientFTP, lcDirRemote, lcDirLocal)
    LOCAL x, lOk, lnSuccess, cmsgErrLog, nQtdeFiles
    LOCAL lcRemoteFileName, lcLocalFileName, loRemoteFiles, loFileFTP, lcNewDirLocal
    LOCAL cDirMes, cDirAno, nTamNew
 
    loFileFTP = CreateObject('Chilkat_9_5_0.FileAccess')
    lOk = .T.
 
    * Abre o diretório remoto
    lcHandle = loSftp.OpenDir(lcDirRemote)
    lnSuccess = loSftp.LastMethodSuccess
 
    * Verifica se o diretório foi acessado com sucesso
    If lOk And lnSuccess <> 1
        lOk = .F.
        c*msgErrLog = loSftp.LastErrorText
        RELEASE loSftp
        Return(lOk)
    EndIf
 
    * Obtém a listagem de arquivos no diretório remoto
    loDirListing = loSftp.ReadDir(lcHandle)
 
    * Verifica se a listagem foi bem-sucedida
    If (loSftp.LastMethodSuccess <> 1) THEN
        c*msgErrLog = loSftp.LastErrorText
        RELEASE loSftp
        Return(c*msgErrLog)
    EndIf
 
    i = 0
    n = loDirListing.NumFilesAndDirs
    *msg(cClientFTP + ": Número de arquivos no diretório remoto: " + str(n))
   
    * Itera sobre cada arquivo encontrado
    For i = 0 To n - 1
        loFileObj = loDirListing.GetFileObject(i)
        lcRemoteFileName = loFileObj.Filename
 
        * Obtém a data de modificação do arquivo
        lcRemoteFileDateTime = loFileObj.LastModifiedTimeStr
 
        **//Analisando ficheiros FTP
        WAIT WINDOW NOWAIT "Baixando ficheiros FTP (" + lcRemoteFileName + ") "
 
        * Verifica se o arquivo termina com ".pdf"
        If LOWER(RIGHT(lcRemoteFileName, 4)) $ ".pdf/.jpg/.jpeg/.png/.tif/" And lcRemoteFileDateTime >= DtoS(date() - 5)
            If fNewFile(cClientFTP, lcDirRemote, lcRemoteFileName)
                * Verifica se o arquivo já existe localmente antes de baixar
                If !File(lcDirLocal + lcRemoteFileName)
                    lnSuccess = loSftp.DownloadFileByName(lcDirRemote + lcRemoteFileName, lcDirLocal + lcRemoteFileName)
 
                    * Verifica se o download foi bem-sucedido
                    If lnSuccess = 1
                        fLogFileOk("ALERTA128", lcDirLocal, lcRemoteFileName)
                    Else
                        cResumoLOG = "Erro ao realizar query from u_impstat512"
                        cMensagemLOG = "Erro ao baixar o arquivo: " + lcRemoteFileName + ". Erro: " + loSftp.LastErrorText
                        fGravaLog( 128, "processarArquivos TFT " ,cResumoLOG ,cMensagemLOG )
                    EndIf
                Else
                    **msg("Arquivo já existe localmente: " + lcRemoteFileName)
                EndIf
            endif
        EndIf
 
        RELEASE loFileObj
    EndFor
 
    RELEASE loSftp
    Return lOk
EndFunc
 
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fAdd()
**//|Autor.....:
**//|Data......: 14 de novembro de 2024, 09:00
**//|Descricao.:
**//|Observacao:
**//+-----------------------------------------------------------------------------------//
* Define a função fAdd para adicionar uma linha ao array
Function fAdd(nRow, arrTarget, v1, v2, v3, v4, v5, v6)
    * A função adiciona os valores passados para a linha nRow
    arrTarget[nRow, 1] = v1
    arrTarget[nRow, 2] = v2
    arrTarget[nRow, 3] = v3
    arrTarget[nRow, 4] = v4
    arrTarget[nRow, 5] = v5
    arrTarget[nRow, 6] = v6
EndFunc
 
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fLogFileOk
**//|Autor.....: Victorugo Larini - larini.victorugo@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.:
**//|Observação:
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fNewFile(cClientFTP,lcDirLocal,lcRemoteFileName)
**-----------------------------------------------------------*
    Local x, cTmpSQL, lOk, nQtdPDFs, cFileOrig, nTamFile, cExtFile, cNomeFic
    cTmpSQL = ''
    lOk = .T.
 
    **//Se arquivo temporário usado, fechar
    If Used("xTmpSQL")
        fecha("xTmpSQL")
    EndIf
 
    **//Monta query para consultar se ficheiro já foi baixado
    TEXT TO cTmpSQL TEXTMERGE NOSHOW
        select count(*) as qtde from u_ftpfileimp (nolock)
        where Upper(nomeftp) = Upper('<<cClientFTP>>')
        and Upper(caminho) = Upper('<<lcDirLocal>>')
        and Upper(ficheiro) = Upper('<<lcRemoteFileName>>')
    ENDTEXT
 
    **//Executa query
    u_sqlexec(cTmpSQL,"xTmpSQL")
 
    **//Atualiza variárel conforme resultado da query
    If xTmpSQL.qtde == 0
        lOk = .T.
    Else
        lOk = .F.
    EndIf
 
    fecha("xTmpSQL")
 
    Return(lOk)
 
EndFunc
 
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fGravaLog
**//|Autor.....: Victorugo Larini
**//|Data......: Junho de 2024
**//|Descricao.:
**//|Observação:
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fGravaLog( m_AlertId, m_AlertStamp ,m_Resumo , m_Mensagem )
**-----------------------------------------------------------*
    Local m_DataIni, m_HoraIni, m_DataFim, m_HoraFim, m_Insert
 
    m_DataIni = dtos( date() )
    m_HoraIni = Time()
    m_DataFim = dtos( date() )
    m_HoraFim = Time()
 
    m_Insert = ""
    TEXT TO m_Insert TEXTMERGE NOSHOW
        INSERT INTO u_LogAlertas ( u_LogAlertasStamp, alstamp, alid, Resumo, iData, iHora, fData, fHora, Mensagem,
                                    ousrinis, ousrdata, ousrhora, usrinis, usrdata, usrhora, marcada )
        VALUES (    (select left(newid(),25))
                    , '<<m_AlertStamp>>'
                    , '<<m_AlertId>>'
                    , '<<m_Resumo>>'
                    , '<<m_DataIni>>'
                    , '<<m_HoraIni>>'  
                    , '<<m_DataFim>>'
                    , '<<m_HoraFim>>'
                    , '<<m_Mensagem>>'
                    , '<<m.m_chinis>>'      
                    , '<<m_DataIni>>'
                    , '<<m_HoraIni>>'
                    , '<<m.m_chinis>>'      
                    , '<<m_DataIni>>'
                    , '<<m_HoraIni>>'
                    , 0
                )
    ENDTEXT
    u_sqlexec( m_Insert )
 
EndFunc
 
 
**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fLogFileOk
**//|Autor.....: Victorugo Larini - larini.victorugo@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.:
**//|Observação:
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fLogFileOk(cClientFTP,lcDirLocal,lcRemoteFileName)
**-----------------------------------------------------------*
    Local cTmpSQL
    cTmpSQL = ''
 
    TEXT TO cTmpSQL TEXTMERGE NOSHOW
        INSERT INTO u_FTPFILEIMP(u_FtpFileImpStamp,nomeftp,caminho,ficheiro,dataimport,ousrinis,ousrdata,ousrhora,usrinis,usrdata,usrhora)
        VALUES
        (
        left(newid(),25),
        '<<cClientFTP>>',
        '<<lcDirLocal>>',
        '<<lcRemoteFilename>>',
        GetDate(),
        '<<m.m_chinis>>',
        Cast(Convert(Varchar(10),GetDate(),111) as DateTime),
            Convert(Varchar(8) ,GetDate(),108),
        '<<m.m_chinis>>',
        Cast(Convert(Varchar(10),GetDate(),111) as datetime),
            Convert(Varchar(8) ,GetDate(),108)
        )
    ENDTEXT
 
    IF !u_sqlexec(cTmpSQL)
        *msg("Erro ao executar a consulta: " + cTmpSQL)
    ENDIF
 
EndFunc