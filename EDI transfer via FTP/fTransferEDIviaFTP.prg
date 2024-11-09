**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fTransferEDIviaFTP()
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: 08 de Julho de 2024, 09:00
**//|Descricao.: POD Importar do FTP = TFD, TBFiles, Lisboa
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**//Declara variaveis
Local lRet

**//Inicializa variaveis
lRet = .T.

**//Restaura Diretorio Padrao
Set Default To alltrim( JUSTPATH ( MAIN_DIR ) )

**//---------------------------------
LOCAL loSftp
fConectaFTP("TFD",@loSftp) && Conectar no FTP da TFD
RELEASE loSftp && Encerrar/finalizar uso do objecto
**//msg("TFD OK")
**//---------------------------------

**//---------------------------------
LOCAL loSftp
fConectaFTP("LISBOA",@loSftp) && Conectar no FTP da LISBOA
RELEASE loSftp && Encerrar/finalizar uso do objecto
**//msg("LISBOA OK")
**//---------------------------------

**//---------------------------------
LOCAL loSftp                                                          
fConectaFTP("TBFILES_COMPRAS",@loSftp) && Conectar no FTP da TBFILES  
RELEASE loSftp && Encerrar/finalizar uso do objecto                   
**//msg("TBFILES_COMPRAS OK")                                         
**//---------------------------------

**//---------------------------------
LOCAL loSftp                                                          
fConectaFTP("TBFILES_DISTRIB",@loSftp) && Conectar no FTP da TBFILES  
RELEASE loSftp && Encerrar/finalizar uso do objecto                   
**//msg("TBFILES_DISTRIB OK")                                         
**//---------------------------------

**//---------------------------------
LOCAL loSftp
fConectaFTP("TBFILES_LUSOCARGO",@loSftp) && Conectar no FTP da TBFILES
RELEASE loSftp && Encerrar/finalizar uso do objecto
**//msg("TBFILES_LUSOCARGO OK")
**//---------------------------------

**//---------------------------------
LOCAL loSftp
fConectaFTP("TBFILES_LCARGO_SUL",@loSftp) && Conectar no FTP da TBFILES
RELEASE loSftp && Encerrar/finalizar uso do objecto
**//msg("TBFILES_LCARGO_SUL OK")
**//---------------------------------

**//---------------------------------
LOCAL loSftp                                                          
fConectaFTP("TBFILES_NACIONAL",@loSftp) && Conectar no FTP da TBFILES 
RELEASE loSftp && Encerrar/finalizar uso do objecto                   
**//msg("TBFILES_NACIONAL OK")                                        
**//---------------------------------

**//---------------------------------
LOCAL loSftp                                                          
fConectaFTP("TBFILES_PALLEX",@loSftp) && Conectar no FTP da TBFILES   
RELEASE loSftp && Encerrar/finalizar uso do objecto                   
**//msg("TBFILES_PALLEX OK")                                          
**//---------------------------------

**//---------------------------------
LOCAL loSftp
fConectaFTP("TBFILES_PORTOCARGO",@loSftp) && Conectar no FTP da TBFILES
RELEASE loSftp && Encerrar/finalizar uso do objecto
**//msg("TBFILES_PORTOCARGO OK")
**//---------------------------------

Return(lRet)


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fConectaFTP
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: Move um arquivo do caminho de origem para o caminho de destino
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fConectaFTP(cClientFTP,loSftp)
**-----------------------------------------------------------*
LOCAL lOk, lcHostname,lnPort,lcUserFtp,lcPassFtp,lcDirRemote,lcDirLocal,lcDirTemp,lcDirNet1,lcDirNet2,lcDirPOD1,lcDirPOD2

**//Define o valor das variaveis conforme CLIENT FTP 
lOk = .T.
Do Case
    Case cClientFTP == "TFD"
        lcHostname  = "hostipserver.ddns.net"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/2019/Entregas/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\TFD\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Scanner\"
        lcDirNet2   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Nacional\Apos 2021\"
        lcDirPOD1   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirPOD2   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Agentes\"
        lcDirPOD3   = "\\192.168.1.7\Lusocargo\Reservas\"

    Case cClientFTP == "LISBOA"
        lcHostname  = "ftp.lusocargo-sul.pt"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/LUSOCARGO/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\LusocargoSulTemp\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Scanner\"
        lcDirNet2   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Entregas feitas pela Lusocargo Sul de carga nossa\Lusocargo\"
        lcDirPOD1   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirPOD2   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Agentes\"

    Case cClientFTP == "TBFILES_COMPRAS"
        lcHostname  = "5.9.153.194"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/compras/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Scanner\Contabilidade\PHC - Sede\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = ""
        lcDirNet2   = ""

    Case cClientFTP == "TBFILES_DISTRIB"
        lcHostname  = "5.9.153.194"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/distribuicao/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Distribuicao (guias dos transportadores)\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = ""
        lcDirNet2   = ""

    Case cClientFTP == "TBFILES_LUSOCARGO"
        lcHostname  = "5.9.153.194"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/lusocargo/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\TBFiles\Lusocargo\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Scanner\"
        lcDirNet2   = ""
        lcDirPOD1   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirPOD2   = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Agentes\"

    Case cClientFTP == "TBFILES_LCARGO_SUL"
        lcHostname  = "5.9.153.194"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/lusocargo_sul/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Carga da LusocargoSul entregue por nos\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = ""
        lcDirNet2   = ""

    Case cClientFTP == "TBFILES_NACIONAL"
        lcHostname  = "5.9.153.194"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/nacional/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Nacional\Apos 2021\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = ""
        lcDirNet2   = ""

    Case cClientFTP == "TBFILES_PALLEX"
        lcHostname  = "5.9.153.194"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/pallex/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Pallex\TBFiles\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = ""
        lcDirNet2   = ""

    Case cClientFTP == "TBFILES_PORTOCARGO"
        lcHostname  = "5.9.153.194"
        lnPort      = 21
        lcUserFtp   = "xxxxxxxxxx"
        lcPassFtp   = "xxxxxxxxxx"

        lcDirRemote = "/portocargo/"
        lcDirLocal  = "\\192.168.1.174\Lusocargo\SERVICOS\POD\Portocargo\"
        lcDirTemp   = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("Tmpfiles\TempPOD\")
        lcDirNet1   = ""
        lcDirNet2   = ""

    Otherwise
        lOk = .F.
EndCase


**//------------------------------------------------
**// TFD
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TFD"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk

        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,lcDirLocal,"5*.pdf",lcDirTemp,lcDirNet1,"")

        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,lcDirNet2,"8*.pdf",lcDirTemp,"","")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Pesquisa referencia POD e renomeia ficheiro para ser enviado para FTP do Agente 
        fRenRefPOD(cClientFTP,lcDirPOD1)

        **//Copiar todos os ficheiros de um Diretorio de um lado para outro
        fCopyDir(cClientFTP, lcDirPOD1, lcDirPOD2)

        **//Enviar para o FTP dos Agentes
        fSendPODparaFTP(cClientFTP,lcDirTemp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

        **//Executar o ID.52 (Pasta SCANNER)
**        fScannerPOD(cClientFTP,lcDirNet1,lcDirPOD3) &&//Faz o processo que o ID.52 do My Cargo fazia

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// LUSOCARGO SUL
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "LISBOA"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"LUSOCA*.pdf",lcDirTemp,lcDirNet1,lcDirNet2)

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Pesquisa referencia POD e renomeia ficheiro para ser enviado para FTP do Agente 
        fRenRefPOD(cClientFTP,lcDirPOD1)

        **//Copiar todos os ficheiros de um Diretorio de um lado para outro
        fCopyDir(cClientFTP, lcDirPOD1, lcDirPOD2)

        **//Enviar para o FTP dos Agentes
        fSendPODparaFTP(cClientFTP,lcDirTemp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

        **//Executar o ID.52 (Pasta SCANNER)
        **//fScannerPOD(cClientFTP,lcDirNet1,lcDirPOD3) &&//Faz o processo que o ID.52 do My Cargo fazia

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// TBFILES_COMPRAS
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TBFILES_COMPRAS"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"*.*",lcDirTemp,lcDirNet1,lcDirNet2)

        **//Apagar ficheiros do FTP
        fDelArqFTP(@loSftp,lcDirRemote,"*.*")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Executar o ID.53,ID.54,ID.55, (Pasta SCANNER/Contabilidade)
        **//fScannerCTB(cClientFTP,lcDirNet1,lcDirPOD3) &&//Faz o processo que o ID.53,ID.54,ID.55 do My Cargo fazia

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// TBFILES_DISTRIBUICAO
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TBFILES_DISTRIB"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"*.*",lcDirTemp,lcDirNet1,lcDirNet2)

        **//Apagar ficheiros do FTP
        fDelArqFTP(@loSftp,lcDirRemote,"*.*")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// TBFILES_LUSOCARGO
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TBFILES_LUSOCARGO"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"*.*",lcDirTemp,lcDirNet1,lcDirNet2)

        **//Pesquisa referencia POD e renomeia ficheiro para ser enviado para FTP do Agente 
        fRenRefPOD(cClientFTP,lcDirPOD1)

        **//Copiar todos os ficheiros de um Diretorio de um lado para outro
        fCopyDir(cClientFTP, lcDirPOD1, lcDirPOD2)

        **//Executar o ID.52 (Pasta SCANNER)
        **//fScannerPOD(cClientFTP,lcDirNet1,lcDirPOD3) &&//Faz o processo que o ID.52 do My Cargo fazia

        **//Apagar ficheiros do FTP
        fDelArqFTP(@loSftp,lcDirRemote,"*.*")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Enviar para o FTP dos Agentes
        fSendPODparaFTP(cClientFTP,lcDirTemp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// TBFILES_LUSOCARGO_SUL
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TBFILES_LCARGO_SUL"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//Cria a pasta MAIA para salvar os ficheiros la dentro e facilitar o upload
        If !Directory(lcDirTemp + 'MAIA\')
            MkDir(lcDirTemp + 'MAIA\')
        EndIf

        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"*.*",lcDirTemp+'MAIA\',lcDirNet1,lcDirNet2)

        **//Apagar ficheiros do FTP
        fDelArqFTP(@loSftp,lcDirRemote,"*.*")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        *//Enviar para FTP lisboa2(MAIA)
        fSendPODparaFTP(cClientFTP,lcDirTemp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// TBFILES_NACIONAL
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TBFILES_NACIONAL"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"*.*",lcDirTemp,lcDirNet1,lcDirNet2)

        **//Apagar ficheiros do FTP
        fDelArqFTP(@loSftp,lcDirRemote,"*.*")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// TBFILES_PALLEX
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TBFILES_PALLEX"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"*.*",lcDirTemp,lcDirNet1,lcDirNet2)

        **//Apagar ficheiros do FTP
        fDelArqFTP(@loSftp,lcDirRemote,"*.*")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

    EndIf
EndIf
**//------------------------------------------------


**//------------------------------------------------
**// TBFILES_PORTOCARGO
**//------------------------------------------------
If lOk And lnPort == 21 And cClientFTP == "TBFILES_PORTOCARGO"
    **//21: Conecta ao FTP
    lOk = fLigaFTP(cClientFTP,@loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
    If lOk
        **//Cria a pasta MAIA para salvar os ficheiros la dentro e facilitar o upload
        If !Directory(lcDirTemp + 'PORTOCARGO\')
            MkDir(lcDirTemp + 'PORTOCARGO\')
        EndIf

        **//21: Download POD via FTP
        fDwn21POD(@loSftp,cClientFTP,lcDirRemote,@lcDirLocal,"*.*",lcDirTemp+'PORTOCARGO\',lcDirNet1,lcDirNet2)

        **//Apagar ficheiros do FTP
        fDelArqFTP(@loSftp,lcDirRemote,"*.*")

        **//21: Desconecta do FTP
        fDesligaFTP(@loSftp)

        **//Enviar para o FTP da PortoCargo (PORTO CARGO)
        fSendPODparaFTP(cClientFTP,lcDirTemp)

        **//Apagar ficheiros da Pasta Temporarios utilizado para manipular os ficheiros
        fDelArqLocal(lcDirTemp)

    EndIf
EndIf
**//------------------------------------------------


**//Restaura Diretorio Padrao
Set Default To alltrim( JUSTPATH ( MAIN_DIR ) )

EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fSendPODparaFTP
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Novembro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fSendPODparaFTP(cClientFTP,cDirPODsAg)
**-----------------------------------------------------------*
Local x, nQtdDIRs, cDirTemp, cSeekAgente
Local oSftp,cHostname,nPort,cUserFtp,cPassFtp,cPathUpload1,cPathUpload2
DIMENSION aArqLocal(1)

**//Total de sub-diretorios
nQtdDIRs = aDir(aArqLocal,cDirPODsAg+"*","D")

**//Percorre todos sub-diretorios encontrados
If nQtdDIRs > 0

    For x=1 To nQtdDIRs
        **//Apaga Diretorios
        If AllTrim(Upper(StrTran(aArqLocal(x, 5),".",""))) == "D" And !(AllTrim(aArqLocal(x, 1)) $ ("\.\..\...\"))
            cDirTemp = AddBS(cDirPODsAg) + AddBS(AllTrim(aArqLocal(x, 1)))

            cSeekAgente = AllTrim(aArqLocal(x, 1))
            If !(Upper(cSeekAgente) $ "/LISBOA/MAIA/PORTOCARGO/")
                cSeekAgente = SubStr(cSeekAgente,1,At("-",cSeekAgente)-1)
            EndIf
            lOk = .T.

            **//Busca informacoes para conectar no FTP
            If lOk
                lOk = fFindAcessFTP(cSeekAgente,@cHostname,@nPort,@cUserFtp,@cPassFtp,@cPathUpload1,@cPathUpload2)
            EndIf

            **//Tenta conectar no FTP
            If lOk
                lOk = fLigaFTP("UPLOADFTP_"+cSeekAgente,@oSftp,cHostname,nPort,cUserFtp,cPassFtp)
            EndIf

            **//Se conectou, envia ficheiros para o FTP
            If lOk
                fUploadFile(cClientFTP,oSftp,cHostname,nPort,cDirTemp,cPathUpload1)
                If !Empty(cPathUpload2)
                    fUploadFile(cClientFTP,oSftp,cHostname,nPort,cDirTemp,cPathUpload2)
                EndIf
            EndIf

            **//Desconect do FTP ou SFTP
            fDesligaFTP(@oSftp)

        EndIf
    Next x
EndIf

Return


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fUploadFile
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Novembro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fUploadFile(cClientFTP,oSftp,cHostname,nPort,cDirTemp,cPathUpload)
**-----------------------------------------------------------*
Local n, nRet, nTotalArquivos
DIMENSION aArquivos(1)

**//Lista arquivos da pasta
nTotalArquivos = aDir(aArquivos, cDirTemp + "*.*")

If nPort == 21
    **//Percorre todos os ficheiros da pasta e envia para o FTP
    For n = 1 To nTotalArquivos
        cArquivo = Alltrim(aArquivos(n,1))

        **//Enviando ficheiros para FTP
        WAIT WINDOW NOWAIT "Enviando ficheiros ("+cArquivo+") para FTP ("+cHostname+")  "

        cFileOrig = cDirTemp+cArquivo
        cFileDest = cPathUpload+cArquivo

        **//Envia para Servidor FTP
        nRet = oSftp.PutFile(cFileOrig,cFileDest)

        **//Enviado com sucesso, entao pode apagar
        If (nRet == 1) AND File(cFileOrig)
            Delete File &cFileOrig.  && Apaga o arquivo
        EndIf
    Next n
EndIf


If nPort == 22
    **//Percorre todos os ficheiros da pasta e envia para o FTP
    For n = 1 To nTotalArquivos
        cArquivo = Alltrim(aArquivos(n,1))

        **//Enviando ficheiros para FTP
        WAIT WINDOW NOWAIT "Enviando ficheiros ("+cArquivo+") para FTP ("+cHostname+")  "

        cFileOrig = cDirTemp+cArquivo
        cFileDest = cPathUpload+cArquivo

        **//Envia para Servidor FTP
        nRet = oSftp.UploadFileByName(cFileDest,cFileOrig)

        **//Enviado com sucesso, entao pode apagar
        If (nRet == 1) AND File(cFileOrig)
            **//Apaga o arquivo
            Delete File &cFileOrig.

            **//Grava LOG dos ficheiro baixado do servidor e saldo com sucesso LOCAL
            fLogFileOk("UPLOAD_"+cClientFTP,cFileOrig,cFileDest)
        Else
            **//Grava LOG
            cMsgErrLog = loSftp.LastErrorText
            cResumoLOG = "ERRO - fUploadFile("+cArquivo+") - Entidade: "+cClientFTP
            cMensagemLOG = "Erro enviar ficheiro para Servidor - Dir.Origem:"+cFileOrig+" - LOG: " + StrTran(cMsgErrLog,Chr(13)," ")
            fGravaLog( 778, "fTransferEDIviaFTP("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )
        EndIf
    Next n
EndIf

**//Fecha a janela de espera apos o termino do processamento
WAIT CLEAR

Return

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fDesligaFTP
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Novembro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fDesligaFTP(oObjFTP)
**-----------------------------------------------------------*
Local lOk
lOk = .T.

Try
    oObjFTP.Disconnect()
    Release oObjFTP
Catch To oErr
    lOk = .F.
EndTry

Return(lOk)


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fFindAcessFTP
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Novembro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fFindAcessFTP(cCodAgente,cHostname,nPort,cUserFtp,cPassFtp,cPathUpload1,cPathUpload2)
**-----------------------------------------------------------*
Local lOk, cTmpQry

lOk          = .F.
cHostname    = ""
nPort        = ""
cUserFtp     = ""
cPassFtp     = ""
cPathUpload1 = ""
cPathUpload2 = ""

cTmpQry = ''
TEXT TO cTmpQry NOSHOW TEXTMERGE
    select 
    '128' as modelo,
    A.a_cod_phc,
    A.a_entidadeEDI,
    Max(A.a_nome) as nome_agente,
    B.servidor_pod,
    B.servidor_pod_ssl,
    B.username_pod,
    B.password_pod,
    B.upload_url_pod
    from [EDI_Lusocargo].[dbo].[Agente] as A
    inner join [EDI_Lusocargo].[dbo].[Credenciais128] as B on B.entidadeEDI = A.a_entidadeEDI
    where A.a_cod_phc = <<cCodAgente>>
    group by A.a_entidadeEDI, A.a_cod_phc, B.servidor_pod_ssl, B.servidor_pod, B.username_pod, B.password_pod, B.upload_url_pod
    
    union all
    
    select
    '512' as modelo,
    A.a_cod_phc,
    A.a_entidadeEDI,
    Max(A.a_nome) as nome_agente,
    B.servidor_pod,
    B.servidor_pod_ssl,
    B.username_pod,
    B.password_pod,
    B.upload_url_pod
    from [EDI_Lusocargo].[dbo].[Agente512] as A
    inner join [EDI_Lusocargo].[dbo].[Credenciais512] as B on B.entidadeEDI = A.a_entidadeEDI
    where A.a_cod_phc = <<cCodAgente>>
    group by A.a_entidadeEDI, A.a_cod_phc, B.servidor_pod_ssl, B.servidor_pod, B.username_pod, B.password_pod, B.upload_url_pod
    
ENDTEXT

**//Se arquivo temporario usado, fechar
If Used("xTmpSQL")
    fecha("xTmpSQL")
EndIf

**//Executa query
If u_sqlexec(cTmpQry,"xTmpSQL")

    **// Seleciona cursor
    Select xTmpSQL
    
    **// Posiciona no primeiro registo do cursor
    Go Top

    **// Verifica se tem registo
    If (!Eof("xTmpSQL"))
        lOk          = .T.
        cHostname    = Alltrim(xTmpSQL.servidor_pod)
        nPort        = xTmpSQL.servidor_pod_ssl
        cUserFtp     = Alltrim(xTmpSQL.username_pod)
        cPassFtp     = Alltrim(xTmpSQL.password_pod)
        cPathUpload1 = Alltrim(xTmpSQL.upload_url_pod)
    EndIf
EndIf

**//Este caso nao tem cadastro
If lOk == .F. And Upper(cCodAgente) == "LISBOA"
    lOk          = .T.
    cHostname    = "193.126.8.166" 
    nPort        = 21
    cUserFtp     = "xxxxxxxxxx"
    cPassFtp     = "xxxxxxxxxx"
    cPathUpload1 = "/" 
EndIf

**//Este caso nao tem cadastro
If lOk == .F. And Upper(cCodAgente) == "MAIA"
    lOk          = .T.
    cHostname    = "ftp.lusocargo-sul.pt" 
    nPort        = 21
    cUserFtp     = "xxxxxxxxxx"
    cPassFtp     = "xxxxxxxxxx"
    cPathUpload1 = "/MAIA/"
    cPathUpload2 = "/MAIA/POD/"
EndIf

**//Este caso nao tem cadastro
If lOk == .F. And Upper(cCodAgente) == "PORTOCARGO"
    lOk          = .T.
    cHostname    = "5.9.153.194" 
    nPort        = 21
    cUserFtp     = "xxxxxxxxxx"
    cPassFtp     = "xxxxxxxxxx"
    cPathUpload1 = "/POD/"
EndIf

**//Se arquivo temporario usado, fechar
If Used("xTmpSQL")
    fecha("xTmpSQL")
EndIf

Return(lOk)

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fDelArqFTP
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Outubro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fDelArqFTP(loSftp,lcDirRemote,lcFilenamePattern)
**-----------------------------------------------------------*
Local x, lnSuccess, nQtdeFiles

lnSuccess = loSftp.ChangeRemoteDir(lcDirRemote)

If lnSuccess == 1

    **//Apagando ficheiros FTP
    WAIT WINDOW NOWAIT "Deletando ficheiros do FTP... "

    loSftp.ListPattern = lcFilenamePattern
    nQtdeFiles = loSftp.GetDirCount()

    For x = 0 To nQtdeFiles - 1
        lOk = .T.
        **//Apenas ficheiros e NAO sub-diretorios
        If loSftp.GetIsDirectory(x) <> 1
            **//Obtenha detalhes do Ficheiro
            lcRemoteFileName = AllTrim(loSftp.GetFilename(x)) && Nome do ficheiro
            loSftp.DeleteRemoteFile(lcRemoteFileName)

            **//Apagando ficheiros FTP
            WAIT WINDOW NOWAIT "Deletando ficheiros do FTP ( "+lcRemoteFileName+" ) "
        EndIf
    Next x
EndIf

Return


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fDelArqLocal
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Setembro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fDelArqLocal(lcDirTemp)
**-----------------------------------------------------------*
Local x, nQtdDIRs, cDelTemp
DIMENSION aArqLocal(1)

**//Total de sub-diretorios
nQtdDIRs = aDir(aArqLocal,lcDirTemp+"*","D")

**//Percorre todos sub-diretorios encontrados
If nQtdDIRs > 0
    **//Apagando ficheiros FTP
    WAIT WINDOW NOWAIT "Deletando ficheiros e pastas LOCAL ... "

    **//Pausa de 3 segundos
    INKEY(3)

    For x=1 To nQtdDIRs
        cDelTemp = ""

        **//Apagando ficheiros FTP
        WAIT WINDOW NOWAIT "Deletando ficheiros e pastas LOCAL ( "+AllTrim(aArqLocal(x, 1))+" ) "

        **//Apaga Ficheiros da pasta RAIZ
        If AllTrim(Upper(StrTran(aArqLocal(x, 5),".",""))) != "D" And !(AllTrim(aArqLocal(x, 1)) $ ("\.\..\...\"))
            cDelTemp = AddBS(lcDirTemp) + AllTrim(aArqLocal(x, 1))
            If File(cDelTemp)
                DELETE FILE &cDelTemp.
            EndIf
        EndIf

        **//Apaga Ficheiros das sub-pastas
        If AllTrim(Upper(StrTran(aArqLocal(x, 5),".",""))) == "D" And !(AllTrim(aArqLocal(x, 1)) $ ("\.\..\...\"))
            cDelTemp = AddBS(lcDirTemp) + AddBS(AllTrim(aArqLocal(x, 1)))
            **// Apaga todos os arquivos da pasta
            ERASE (cDelTemp + "*.*")
        EndIf

        **//Apaga Diretorios
        If AllTrim(Upper(StrTran(aArqLocal(x, 5),".",""))) == "D" And !(AllTrim(aArqLocal(x, 1)) $ ("\.\..\...\"))
            cDelTemp = AddBS(lcDirTemp) + AddBS(AllTrim(aArqLocal(x, 1)))
            **// Apaga pasta
            If Directory(cDelTemp)
                RmDir(cDelTemp)
            EndIf
        EndIf

    Next x

    **//Fecha a janela de espera apos o termino do processamento
    WAIT CLEAR

EndIf

Release aArqLocal

EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fRenRefPOD
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Setembro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fRenRefPOD(cClientFTP,cDirPOD)
**-----------------------------------------------------------*
Local x, nQtdPDFs, cDirDest, cTmpQry, cFileOrig, cFileName, cNewFileName, cFileNameApagar, cDirLisboa, cExtFile, cNumGuia, nPosChr

DIMENSION aArquivos(1)

**//Apagando ficheiros FTP
WAIT WINDOW NOWAIT "Renomeando ficheiros LOCAL... "

nQtdPDFs = aDir(aArquivos,cDirPOD+"5*.pdf")
For x=1 To nQtdPDFs
    cFileOrig = Alltrim(Upper(aArquivos(x,1)))

    **//Apagando ficheiros FTP
    WAIT WINDOW NOWAIT "Renomeando ficheiros LOCAL ( "+cFileOrig+" ) "

    cExtFile = Right(cFileOrig,4) 
    cNumGuia = StrTran(cFileOrig,cExtFile,"")

    nPosChr = At("_",cNumGuia)-1
    If nPosChr >= 5
        cNumGuia = SubStr(cFileOrig,1,nPosChr)
    EndIf

	cTmpQry = ''
    TEXT TO cTmpQry NOSHOW TEXTMERGE

        select 
        bo.no,
        'lusocargo_entrega' = bo.obrano, 
        'reserva' = bo2.u_resno,
        'agente_ref' = (select replace(replace(replace(replace(replace(replace(replace(replace(replace(YY.u_factobs,' ',''),'/',''),'\',''),'.',''),'-',''),'(',''),')',''),',',''),'"','') from bo as XX (nolock), bo2 as YY (nolock) where XX.ndos=4 and XX.obrano=bo2.u_resno and YY.bo2stamp=XX.bostamp) ,
        'agente_cod' = (select YY.agenteref from bo as XX (nolock), u_bo3 as YY (nolock) where XX.ndos=4 and XX.obrano=bo2.u_resno and YY.bostamp=XX.bostamp),
        'agente_nome' = (select replace(replace(replace(replace(replace(replace(replace(replace(replace(YY.agentenome,' ',''),'/',''),'\',''),'.',''),'-',''),'(',''),')',''),',',''),'"','')  from bo as XX (nolock), u_bo3 as YY (nolock) where XX.ndos=4 and XX.obrano=bo2.u_resno and YY.bostamp=XX.bostamp)
        from bo(nolock)
        left join bo2(nolock) on bo.bostamp=bo2.bo2stamp
        where bo.ndos=11
        and bo.obrano='<<cNumGuia>>'
    ENDTEXT

    **//Se arquivo temporario usado, fechar
    If Used("xTmpSQL")
        fecha("xTmpSQL")
    EndIf

    If !u_sqlexec(cTmpQry,"xTmpSQL")
        MENSAGEM("ERRO na leitura da guia. P.f. contacte a Informatica."+Chr(13)+Chr(10)+"Ficheiro: "+cFileOrig,"DIRECTA")
        msg(cTmpQry)
    Else

        cNewFileName = Alltrim(xTmpSQL.agente_ref)
        cFileNameApagar = Alltrim(cDirPOD) + Alltrim(cFileOrig)
        cDirLisboa = ''

        **//Se nao tem referencia, entao usar o numero da nossa reserva
        If Empty(cNewFileName)
            cNewFileName = "RESERVA_LC_"+aStr(xTmpSQL.reserva)
        EndIf

        If xTmpSQL.no=2287 And aStr(xTmpSQL.agente_cod) $ '259616/257306/257289/257312/257290/257284/257307/257286/257308/257288/259617/259618/257291/259619/257311/257287/' 
            If !Directory(cDirPOD + 'LISBOA\')
                MkDir(cDirPOD + 'LISBOA\')
            EndIf

            cDirLisboa = cDirPOD + 'LISBOA\' + alltrim(cNewFileName)+'.pdf'
            Do Case
                Case "TBFILES" $ cClientFTP
                    cNewFileName = "Este_Nao"
                
                Case "TBFILES_LUSOCARGO" $ cClientFTP
                    cNewFileName = "Este_Nao"+alltrim(cNewFileName)+'.pdf'

                Otherwise
                    cNewFileName = alltrim(cNewFileName)+'.pdf'
            EndCase
        Else
            cNewFileName = alltrim(cNewFileName)+'.pdf'
        Endif

        **//Verifique se o diretorio de destino existe, caso contrario, crie-o
        cDirDest = cDirPOD + aStr(xTmpSQL.agente_cod) +"-"+ Upper(Alltrim(PadR(xTmpSQL.agente_nome,11))) + "\"

        If !Directory(cDirDest)
            MkDir(cDirDest)
        EndIf

        If File(cDirDest + cNewFileName)
            cExtFile = Right(cNewFileName,4)    
            cNewFileName = StrTran(cNewFileName,cExtFile,"_"+DtoS(Date())+"_"+StrTran(Time(),":","") + cExtFile)
        EndIf

        If cClientFTP == "TFD" And !Empty(cDirLisboa)
            If File(cDirLisboa)
                **//Renomeia ficheiro
                cExtFile = Right(cDirLisboa,4)
                cDirLisboa = StrTran(cDirLisboa,cExtFile,"_"+DtoS(Date())+"_"+StrTran(Time(),":","") + cExtFile)
            EndIf
            COPY FILE (cFileNameApagar) TO (cDirLisboa)
        EndIf

        Do Case
            Case "Este_Nao" == cNewFileName
                **//Nao fazer Nada
            Case "Este_Nao" $ cNewFileName
                COPY FILE (cFileNameApagar) TO (cDirDest + cNewFileName)
            Otherwise
                COPY FILE (cFileNameApagar) TO (cDirDest + cNewFileName)
        EndCase

        DELETE FILE &cFileNameApagar.

    EndIf

Next x

Release aArquivos

EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fScannerPOD
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fScannerPOD(cClientFTP,lcDirLocal,cRedePOD)
**-----------------------------------------------------------*
Local nRegistos, cDirPOD, cFilePOD, cNewFilePOD, cStatusPOD, cGuiaEntrega, cResumoLOG, cMensagemLOG, cTmpSQL

DIMENSION aFilesPDF(1)

nRegistos = aDir(aFilesPDF, lcDirLocal+"*.pdf")

* inicia a regua
REGUA(0, nRegistos, "Importando POD para Gestao Doc.")

For i = 1 To nRegistos
    cDirPOD = aFilesPDF(i,1)
    cFilePOD = JustFName(cDirPOD)
    
    cGuiaEntrega = ""
    cStatusPOD = "35"

    If At("_", cFilePOD) > 0
        cGuiaEntrega = Left(cFilePOD, At("_", cFilePOD) - 1)
        If !fIsDigit(cGuiaEntrega)
            texto = IdTarefa + "Importacao de POD: Nome de ficheiro invalido."

            **//Grava LOG
            cResumoLOG = "ERRO - fScannerPOD("+cGuiaEntrega+") - Nome de ficheiro invalido."
            cMensagemLOG = "Erro ao validar ficheiro - Dir:"+lcDirLocal+" - File: " + cFilePOD
            fGravaLog( 778, "fTransferEDIviaFTP("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )

            Loop
        EndIf
        **//cStatusPOD = SubStr(cFilePOD, At("_", cFilePOD) + 1, At(".", cFilePOD) - At("_", cFilePOD) - 1)
    Else
        cGuiaEntrega = Left(cFilePOD, At(".", cFilePOD) - 1)
        If !fIsDigit(cGuiaEntrega)

            **//Grava LOG
            cResumoLOG = "ERRO - fScannerPOD("+cGuiaEntrega+") - Nome de ficheiro invalido."
            cMensagemLOG = "Erro ao validar ficheiro - Dir:"+lcDirLocal+" - File: " + cFilePOD
            fGravaLog( 778, "fTransferEDIviaFTP("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )

            Loop
        EndIf
    EndIf

    **//Muda nome do ficheiro
    cNewFilePOD = "POD_" + cFilePOD

    **//Consulta a Guia no PHC
    cTmpSQL = ''
    TEXT TO cTmpSQL TEXTMERGE NOSHOW
        SELECT  BO2.U_RESNO as reserva, 
                bo2.u_resstamp as u_resstamp, 
                bo.no as cliente, 
                BO2.U_PRONO as processo,
                (SELECT TOP 1 bo2.u_factobs FROM bo (nolock) WHERE bo.bostamp = bo2.u_resstamp) as ref_agente,
                (SELECT TOP 1 u_bo3.agenteref as cod_agente FROM u_bo3 (nolock) WHERE u_bo3.bostamp = bo2.u_resstamp) as cod_agente,
                (SELECT TOP 1 bor.edi FROM bo (nolock) bor WHERE bor.ndos=4 AND bor.obrano = BO2.U_RESNO) as edi
         FROM bo (nolock)
        INNER JOIN bo2 (nolock) ON bo2.bo2stamp=bo.bostamp
        WHERE bo.ndos = 11 
          AND bo.obrano = <<cGuiaEntrega>> 
          AND LTRIM(RTRIM(BO2.U_VIA)) = 'R'
    ENDTEXT

    **//Executa query
    u_sqlexec(cTmpSQL,"TempSQL")

    **//Verifica se query retorno resultado
    If Used("TempSQL") And Reccount("TempSQL") > 0

        Do While (!Eof("TempSQL"))
            lOk = .T.
            
            **//Verifica de o numero da reserva e processo estao preenchidos
            If lOk And (Empty(TempSQL.reserva) Or Empty(TempSQL.processo))
                lOk = .F.
                **//Tratar MSG erro
                msg("Guia de Entrega sem Processo ou Reserva associada")    
            EndIf
            
            **//Cria diretorio caso nao exista
            If lOk And !Directory(cRedePOD + aStr(TempSQL.reserva))
                **//Tenta criar o diretorio de destino
                Try
                    MkDir(cRedePOD + aStr(TempSQL.reserva))
                Catch To oErr
                    lOk = .F.
                    msg("Erro ao criar o diretorio ("+cRedePOD + aStr(TempSQL.reserva)+")")
                EndTry
            EndIf
            
            **//Verifica se ficheiro ja existe no servidor
            If lOk And File(cRedePOD + aStr(TempSQL.reserva) + "\" + cNewFilePOD)
                lOk = .F.

                **//Verifica tamanho do ficheiro, se for do mesmo tamanho, entao mantem os dois ficheiros
                lLoop = .T.
                nTamFile = filesize(lcDirLocal+cFilePOD)
                nTamNew = filesize(cRedePOD + aStr(TempSQL.reserva) + "\" + cNewFilePOD)
                nCount = 1
                cCount = "1"
                Do While lLoop And nTamFile != nTamNew
                    nCount = nCount + 1
                    cCount = StrZero(nCount,3,0)
                    cNewFilePOD = StrTran("POD_" + cFilePOD,".PDF","_"+cCount+".PDF")

                    **//Se ficheiro existe, verifica o tamanho 
                    If File(cRedePOD + aStr(TempSQL.reserva) + "\" + cNewFilePOD)
                        nTamNew = filesize(cRedePOD + aStr(TempSQL.reserva) + "\" + cNewFilePOD)
                        **//Se tamanho igual, entao nao gravar novamente
                        If nTamFile == nTamNew
                            lLoop = .F.
                        EndIf
                    Else
                        **//Se nao encontrou, entao sai do LOOP e grava
                        lLoop = .F.
                        lOk = .T.
                    EndIf

                EndDo

                If !lOk
                    **//Tratar MSG erro
                    msg("POD da Guia de Entrega ja importado")
                EndIf
            EndIf

            If lOk
                * Retornar tamanho do ficheiro
                nTamFile = filesize(lcDirLocal+cFilePOD)

                * Copia Ficheiro para rede
                COPY FILE (lcDirLocal + cFilePOD) TO (cRedePOD + aStr(TempSQL.reserva) + "\" + cNewFilePOD)

                * Insert into MySQL
                cIncPath         = cRedePOD + aStr(TempSQL.reserva) + "\" + cNewFilePOD
                cIncOwner        = "pod"
                cIncMime_type    = "application/octet-stream"
                nIncTamanho      = nTamFile
                cIncPrivado      = "0"
                cIndDataHora     = fGetDtHrFile(lcDirLocal + cFilePOD)
                cIncCliente      = TempSQL.cliente
                cIncProcesso     = TempSQL.processo
                cIncReserva      = TempSQL.reserva
                cIncAgente       = TempSQL.cod_agente
                cIncGuia_entrega = cGuiaEntrega
                cIncRef_agente   = Alltrim(TempSQL.ref_agente)
                cIncReserva_rod  = TempSQL.reserva

                **MyCmd = "INSERT INTO ficheiro (path, owner, mime_type, tamanho, privado, cliente, processo, reserva, agente, guia_entrega, ref_agente, reserva_rod) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                **//SQLEXEC(MyConn, MyCmd, remote_path, "pod", "application/octet-stream", nTamFile, "0", cliente, processo, reserva, cod_agente, guia_entrega, ref_agente, reserva)

                **//--------------------------------
                **// INSERT mycargo_lcargo.ficheiro
                **//--------------------------------
                cTempSQL = ''
                TEXT TO cTempSQL TEXTMERGE NOSHOW

                    INSERT INTO OPENQUERY(MYSQLMYCARGO, 
                    'SELECT 
                        path, 
                        owner, 
                        mime_type, 
                        tamanho, 
                        privado, 
                        datahora,
                        cliente, 
                        processo, 
                        reserva, 
                        agente, 
                        guia_entrega, 
                        ref_agente, 
                        reserva_rod
                    FROM mycargo_lcargo.ficheiro'
                    )

                    VALUES 
                    (   '<<cIncPath>>', 
                        '<<cIncOwner>>', 
                        '<<cIncMime_type>>', 
                        <<nIncTamanho>>, 
                        '<<cIncPrivado>>', 
                        '<<cIndDataHora>>'
                        '<<cIncCliente>>', 
                        '<<cIncProcesso>>', 
                        '<<cIncReserva>>', 
                        '<<cIncAgente>>', 
                        '<<cIncGuia_entrega>>', 
                        '<<cIncRef_agente>>', 
                        '<<cIncReserva_rod>>'
                    );

                ENDTEXT
                **//Executa query
                msg(cTempSQL)
                *u_sqlexec(cTempSQL)


                **//--------------------------------
                **// UPDATE u_bo3
                **//--------------------------------
                cTempSQL = ''
                TEXT TO cTempSQL TEXTMERGE NOSHOW
                    UPDATE u_bo3 SET e = '<<cStatusPOD>>' WHERE u_bo3.bostamp = (SELECT TOP 1 bostamp FROM bo WHERE obrano = '<<cIncReserva>>' AND ndos = 4)
                ENDTEXT
                **//Update PHC
                msg(cTempSQL)
                *u_sqlexec(cTempSQL)


                **//--------------------------------
                **// INSERT u_track
                **//--------------------------------
                cTempSQL = ''
                u_stamp=u_stamp()
                TEXT TO cTempSQL TEXTMERGE NOSHOW
                    INSERT INTO u_track (u_trackstamp, no, data, reg, estado, stampreserva, ousrinis, ousrdata, usrinis, usrdata, dataevento) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ENDTEXT
                **//Insert PHC
                datahora = TTOC(DATETIME())
                cmd = ""
                **//SQLEXEC(con, cmd, PHC.RetornaNovoStamp(), reserva, datahora, "I", status, stamp_reserva, "Adm", datahora, "Adm", datahora, datahora)
            EndIf

            SKIP IN TempSQL
        EndDo
    Else
        **//Tratar MSG erro
        msg("Guia de Entrega nao encontrada no PHC") 
    EndIf

Next i

* fecha a regua
REGUA(2)

Release aFilesPDF

EndFunc


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fGetDtHrFile
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Outubro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fGetDtHrFile(cArquivo)
**-----------------------------------------------------------*
LOCAL cDataHora
DIMENSION aArquivo(1)
cDataHora = ""

* Usa ADIR para obter as informacoes do arquivo
If aDir(aArquivo, cArquivo) > 0
    * A terceira coluna contem a data, e a quarta coluna contem a hora
    cDataHora = DtoC(aArquivo[3]) + " " + TtoC(aArquivo[4], 1)  && Formata data e hora
EndIf

Return(cDataHora)
EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fIsDigit
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fIsDigit(cString)
**-----------------------------------------------------------*
Local i, lIsDigit
lIsDigit = .T.  && Assume que a string e numerica

For i = 1 To Len(cString)
    If Not Between(SubStr(cString, i, 1), "0", "9")
        lIsDigit = .F.  && Encontrou um caractere nao numerico
        Exit
    EndIf
EndFor

Return(lIsDigit)
EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fDwn21POD
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fDwn21POD(loSftp,cClientFTP,lcDirRemote,lcDirLocal,lcFilenamePattern,lcDirCopiaDest1,lcDirCopiaDest2,lcDirCopiaDest3)
**-----------------------------------------------------------*
LOCAL x, lOk, lnSuccess, cMsgErrLog, nQtdeFiles
LOCAL lcRemoteFileName, lcLocalFileName, loRemoteFiles, loFileFTP, lcNewDirLocal
Local cDirDia, cDirMes, cDirAno

loFileFTP = CreateObject('Chilkat_9_5_0.FileAccess')

lOk = .T.
lnSuccess = loSftp.ChangeRemoteDir(lcDirRemote)

**//Caso ERRO
If lOk And lnSuccess <> 1
    lOk = .F.
    cMsgErrLog = loSftp.LastErrorText
    RELEASE loSftp

    **//Grava LOG
    cResumoLOG = "ERRO - fDwn21POD("+lcFilenamePattern+") - Entidade: "+cClientFTP
    cMensagemLOG = "Erro aceder diretorio no Servidor - Dir:"+lcDirRemote+" - LOG: " + StrTran(cMsgErrLog,Chr(13)," ")
    fGravaLog( 778, "fTransferEDIviaFTP("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )

    Return(lOk)
EndIf

If lOk

    **//Analisando ficheiros FTP
    WAIT WINDOW NOWAIT "Analisando ficheiros FTP... "

    loSftp.ListPattern = lcFilenamePattern
    nQtdeFiles = loSftp.GetDirCount()

    For x = 0 To nQtdeFiles - 1
        lOk = .T.
        **//Apenas ficheiros e NAO sub-diretorios
        If loSftp.GetIsDirectory(x) <> 1
            **//Obtenha detalhes do Ficheiro
            lcRemoteFileName = AllTrim(loSftp.GetFilename(x)) && Nome do ficheiro
            loDt = loSftp.GetLastModDt(x)
            lcRemoteFileDateTime = StrTran(SubStr(loDt.GetAsTimestamp(0),1,10),"-","")

            **//Analisando ficheiros FTP
            WAIT WINDOW NOWAIT "Analisando ficheiros FTP ("+lcRemoteFileName+") "

            If lcRemoteFileDateTime >= DtoS(date()-8)
                **//Define a estrutura que sera salvo o ficheiro (DIR+\+ANO+\+MES)
                cDirAno = SubStr(lcRemoteFileDateTime,1,4)
                cDirMes = SubStr(lcRemoteFileDateTime,5,2)
                cDirDia = SubStr(lcRemoteFileDateTime,7,2)

                **//Se TBFile, pasta compras, nao criar subdiretorio com ANO + MES
                Do Case
                    Case "/"+cClientFTP+"/" $ "/TFD/TBFILES_COMPRAS/TBFILES_LCARGO_SUL/TBFILES_NACIONAL/TBFILES_PORTOCARGO/"
                        lcNewDirLocal = lcDirLocal

                    Case "/"+cClientFTP+"/" $ "/TBFILES_PALLEX/"
                        cMesExt = fNomeMes(cDirMes)
                        lcNewDirLocal = lcDirLocal + cDirAno + "\" + cMesExt + " " + cDirAno + "\" + "Dia " + cDirDia + "\"

                    Otherwise
                        lcNewDirLocal = lcDirLocal + cDirAno + " - " +  cDirMes + "\"
                EndCase

                **// Verifica se o diretorio de destino existe
                If !Directory(lcNewDirLocal)
                    **// Tenta criar o diretorio de destino
                    Try
                        MkDir(lcNewDirLocal)
                    Catch To oErr
                        lOk = .F.
                        **//RELEASE loSftp

                        **//Grava LOG
                        cResumoLOG = "ERRO - fDwn21POD("+lcRemoteFileName+") - Entidade: "+cClientFTP
                        cMensagemLOG = "Erro ao criar o diretorio de destino ("+lcNewDirLocal+"): " + oErr.Message
                        fGravaLog( 778, "fTransferEDIviaFTP("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )
                        Return(lOk)
                    EndTry
                Endif

                lcLocalFileName = lcRemoteFileName

                **//Retira o nome "LUSOCA" do ficheiro
                If cClientFTP == "LISBOA"
                    lcLocalFileName = StrTran(lcRemoteFileName,"LUSOCA","")
                EndIf

                ** Verificar se existe registo ta tabela de log (u_ftpfileimp)
                ** Verificar se ficheiro existe nos nossos servidores
                If lOk
                    nTamNew = loSftp.GetSize(x)
                    If fNewFile(cClientFTP,lcNewDirLocal,lcLocalFileName,nTamNew)
                        lOk = .T.
                        If File(lcNewDirLocal+lcLocalFileName)
                            **//Renomeia ficheiro
                            cExtFile = Right(lcLocalFileName,4)
                            lcLocalFileName = StrTran(lcLocalFileName,cExtFile,"_"+DtoS(Date())+"_"+StrTran(Time(),":","") + cExtFile)
                        EndIf
                    Else
                        lOk = .F.
                    EndIf
                EndIf

                **Baixar somente ficheiros que nao foram baixados anteriormente
                **Baixar apenas novos arquivos
                If lOk
                    Try
                        **Download a file.
                        lnSuccess = loSftp.GetFile(lcDirRemote+lcRemoteFileName,lcNewDirLocal+lcLocalFileName)
                    Catch To oErr
                        lnSuccess = 9999
                    EndTry

                    If lnSuccess <> 1
                        lOk = .F.
                        If lnSuccess = 9999
                            cMsgErrLog = oErr.Message
                        Else
                            cMsgErrLog = loSftp.LastErrorText
                        EndIf
                        **//RELEASE loSftp

                        **//Grava LOG
                        cResumoLOG = "ERRO - fDwn21POD("+lcRemoteFileName+") - Entidade: "+cClientFTP
                        cMensagemLOG = "Erro baixar ficheiro do Servidor - Dir.Orig.:"+lcDirRemote+" - LOG: " + StrTran(cMsgErrLog,Chr(13)," ")
                        fGravaLog( 778, "fTransferEDIviaFTP("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )
                    Else

                        **//Grava LOG dos ficheiro baixado do servidor e saldo com sucesso LOCAL
                        fLogFileOk(cClientFTP,lcNewDirLocal,lcRemoteFileName)

                        **//Faz copia do ficheiro para um segundo local da Rede (Destino 1)
                        If !Empty(lcDirCopiaDest1)
                            fCopyFile(cClientFTP, lcNewDirLocal, lcDirCopiaDest1, lcLocalFileName, lcLocalFileName)
                        EndIf

                        **//Faz copia do ficheiro para um segundo local da Rede (Destino 2)
                        If !Empty(lcDirCopiaDest2)
                            fCopyFile(cClientFTP, lcNewDirLocal, lcDirCopiaDest2, lcLocalFileName, lcLocalFileName)
                        EndIf

                        **//Faz copia do ficheiro para um segundo local da Rede (Destino 3)
                        If !Empty(lcDirCopiaDest3)
                            fCopyFile(cClientFTP, lcNewDirLocal, lcDirCopiaDest3, lcLocalFileName, lcLocalFileName)
                        EndIf

                    EndIf
                Else
                    **//Grava LOG dos ficheiro baixado do servidor e saldo com sucesso LOCAL
                    fLogFileOk(cClientFTP,lcNewDirLocal,lcRemoteFileName)
                EndIf
            EndIf
        EndIf
    Next x

    **//Fecha a janela de espera apos o termino do processamento
    WAIT CLEAR

EndIf

EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fNomeMes
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Outubro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fNomeMes(cMes)
**-----------------------------------------------------------*
LOCAL nMes, aMeses
DIMENSION aMeses(12)

nMes = Int(Val(cMes))

* Define os nomes dos meses em um array
aMeses[01] = "Jan"
aMeses[02] = "Fev"
aMeses[03] = "Mar"
aMeses[04] = "Abr"
aMeses[05] = "Mai"
aMeses[06] = "Jun"
aMeses[07] = "Jul"
aMeses[08] = "Ago"
aMeses[09] = "Set"
aMeses[10] = "Out"
aMeses[11] = "Nov"
aMeses[12] = "Dez"

* Verifica se o numero esta no intervalo correto
If nMes >= 1 AND nMes <= 12
    Return aMeses(nMes)
Else
    Return cMes
EndIf

EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fCopyFile
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fCopyFile(cClientFTP, cDirOrig, cDirDest, cFileNameOrig, cFileNameDest)
**-----------------------------------------------------------*
Local cFileOrigPath, cFileDestPath, lRet, nTamOrig, nTamDest

**//Construa o caminho completo do arquivo de origem e de destino
cFileOrigPath = AddBS(cDirOrig) + Alltrim(cFileNameOrig)
cFileDestPath = AddBS(cDirDest) + Alltrim(cFileNameDest)

**//Verifique se o diretorio de destino existe, caso contrario, crie-o
If !Directory(cDirDest)
    MkDir(cDirDest)
EndIf

**//Copiar o arquivo
Do Case
    **//Se existe na origem e nao existe no distino, entao copia
    Case File(cFileOrigPath) And !File(cFileDestPath)
        Copy File (cFileOrigPath) TO (cFileDestPath)
        fLogFileOk(cClientFTP,cDirDest,cFileNameOrig)
        lRet = .T.

    **//Se existe na origem e tambem existe no distino, entao copia mas renomeria antes para nao sobrepor
    Case File(cFileOrigPath) And File(cFileDestPath)
        **//Renomeia ficheiro
        cExtFile = Right(cFileDestPath,4)
        cFileDestPath = StrTran(cFileDestPath,cExtFile,"_"+DtoS(Date())+"_"+StrTran(Time(),":","") + cExtFile)

        Copy File (cFileOrigPath) TO (cFileDestPath)
        fLogFileOk(cClientFTP,cDirDest,cFileNameOrig)
        lRet = .F.

    Otherwise
        **//Grava LOG
        cResumoLOG = "ERRO - fCopyFile() - File: "+cFileNameOrig
        cMensagemLOG = "Erro ao copiar ficheiro para SCANNER - Dir.Orig.:"+cDirOrig+" - Dir.Dest.: " + cDirDest
        fGravaLog( 778, "fTransferEDIviaFTP("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )
        lRet = .F.
EndCase

Return(lRet)

EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fCopyDir
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Outubro de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fCopyDir(cClientFTP, cDirOrig, cDirDest)
**-----------------------------------------------------------*
Local x, cDirTempOrig, cDirTempDest, nQtdDIRs

DIMENSION aPastas(1)

**//Total de sub-diretorios
nQtdDIRs = aDir(aPastas,cDirOrig+"*","D")

**//Percorre todos sub-diretorios encontrados
If nQtdDIRs > 0
    For x=1 To nQtdDIRs
        If AllTrim(Upper(StrTran(aPastas(x, 5),".",""))) == "D" And !(AllTrim(aPastas(x, 1)) $ ("\.\..\...\"))
            **//Declara variaveis
            cDirTempOrig = AddBS(cDirOrig) + AddBS(AllTrim(aPastas(x, 1)))
            cDirTempDest = AddBS(cDirDest) + AddBS(AllTrim(aPastas(x, 1)))
            fProcFiles(cClientFTP,cDirTempOrig,cDirTempDest)
        EndIf
    Next x
EndIf

Release aPastas

Return


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fProcFiles
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fProcFiles(cClientFTP,cDirTempOrig,cDirTempDest)
**-----------------------------------------------------------*
Local x, nQtdARQs, cFileTemp, nFileSize

DIMENSION aItens(1)

nQtdARQs = aDir(aItens,cDirTempOrig+"*.*")

**//Verifique se o diretorio de destino existe, caso contrario, crie-o
If !Directory(cDirTempDest)
    MkDir(cDirTempDest)
EndIf

**//Apagando ficheiros FTP
WAIT WINDOW NOWAIT "Copiando ficheiros LOCAL... "

**//Copia arquivos para para diretorio destino
For x=1 To nQtdARQs
    cFileTemp = Alltrim(Upper(aItens(x,1)))
    nFileSize = filesize(cDirTempOrig+cFileTemp)

    **//Apagando ficheiros FTP
    WAIT WINDOW NOWAIT "Copiando ficheiros LOCAL ( "+cFileTemp+" ) "

    If fNewFile(cClientFTP, cDirTempDest, cFileTemp, nFileSize)
        fCopyFile(cClientFTP, cDirTempOrig, cDirTempDest, cFileTemp, cFileTemp)
    EndIf
Next x

Release aItens

Return

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fLogFileOk
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fLogFileOk(cClientFTP,lcDirLocal,lcRemoteFileName)
**-----------------------------------------------------------*
Local cTmpSQL, nTamFile
cTmpSQL = ''
nTamFile = filesize(lcDirLocal+lcRemoteFileName)

TEXT TO cTmpSQL TEXTMERGE NOSHOW
    INSERT INTO u_FTPFILEIMP(u_FtpFileImpStamp,nomeftp,caminho,ficheiro,dataimport,ousrinis,ousrdata,ousrhora,usrinis,usrdata,usrhora,size)
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
         Convert(Varchar(8) ,GetDate(),108),
    <<nTamFile>>
    )
ENDTEXT

**//Executa query
u_sqlexec(cTmpSQL)

EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fNewFile
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fNewFile(cClientFTP,lcDirLocal,lcRemoteFileName,nTamNew)
**-----------------------------------------------------------*
Local x, cTmpSQL, lOk, nQtdPDFs, cFileOrig, nTamFile, cExtFile, cNomeFic
cTmpSQL = ''
lOk = .T.

DIMENSION aFicheiros(1)

**//Se arquivo temporario usado, fechar
If Used("xTmpSQL")
    fecha("xTmpSQL")
EndIf

**//Monta query para consultar se ficheiro ja foi baixado
TEXT TO cTmpSQL TEXTMERGE NOSHOW
    select count(*) as qtde from u_ftpfileimp (nolock) 
    where Upper(nomeftp) = Upper('<<cClientFTP>>')
    and Upper(caminho) = Upper('<<lcDirLocal>>')
    and Upper(ficheiro) = Upper('<<lcRemoteFileName>>')
    and (size = <<nTamNew>> or size = 0)
ENDTEXT

**//Executa query
u_sqlexec(cTmpSQL,"xTmpSQL")

**//Atualiza variarel conforme resultado da query
If xTmpSQL.qtde == 0
    lOk = .T.
Else
    lOk = .F.
EndIf
**//Verifca se ficheiro existe no nosso servidor LOCAL
**//Se existir, verifica o tamanho do ficheiro e se for o mesmo, entao nao baixar novamente
If File(lcDirLocal+lcRemoteFileName)

    cExtFile = Right(lcRemoteFileName,4)
    cNomeFic = StrTran(lcRemoteFileName,cExtFile,"")

    nPosChr = At("_",cNomeFic)-1
    If nPosChr >= 5
        cNomeFic = SubStr(cNomeFic,1,nPosChr)
    EndIf

    nQtdPDFs = aDir(aFicheiros,lcDirLocal+cNomeFic+"*.*")
    
    If nQtdPDFs > 0
        lOk = .T.
        For x=1 To nQtdPDFs
            cFileOrig = Alltrim(Upper(aFicheiros(x,1)))
            nTamFile = filesize(lcDirLocal+cFileOrig)

            If nTamFile == nTamNew
                lOk = .F.
            EndIf
        Next x
    EndIf
EndIf

**//Se arquivo temporario usado, fechar
If Used("xTmpSQL")
    fecha("xTmpSQL")
EndIf

Release aFicheiros

Return(lOk)

EndFunc


**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fLigaFTP
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Julho de 2024
**//|Descricao.: 
**//|Observacao: 
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
Function fLigaFTP(cClientFTP,loSftp,lcHostname,lnPort,lcUserFtp,lcPassFtp)
**-----------------------------------------------------------*

lcHostname = AllTrim(lcHostname)
lcUserFtp  = AllTrim(lcUserFtp)
lcPassFtp  = AllTrim(lcPassFtp)

***********************
** FTP
***********************
If lnPort == 21
    LOCAL lnSuccess, cMsgErrLog, lOk
    LOCAL cResumoLOG, cMensagemLOG

    **//Ligacao FTP
    loSftp = CreateObject('Chilkat_9_5_0.Ftp2')

    **//unlock code do PHC para CHILKAT
    lnSuccess = loSftp.unlockcomponent("PHCPToFTP_LsZMMBcM6rwX")

    **//Inicializa variavel
    cMsgErrLog = "ERRO"
    lOk = .F.

    If Type("loSftp")="O"
        **//Atualiza variavel
        cMsgErrLog = "OK"
        lOk = .T.

        **//Define as informacoes do servidor
        loSftp.HostName = lcHostname
        loSftp.UserName = lcUserFtp
        loSftp.Password = lcPassFtp

        **//Conecte-se ao servidor
        lnSuccess = loSftp.Connect()
        WAIT WINDOW "" Timeout 0.2
    EndIf

    **//Caso ERRO
    If lOk And lnSuccess <> 1
        lOk = .F.
        cMsgErrLog = loSftp.LastErrorText
        RELEASE loSftp

        **//Grava LOG
        cResumoLOG = "ERRO - fLigaFTP() - Server: "+lcHostname+" - Port:"+aStr(lnPort)
        cMensagemLOG = "Erro ao autenticar no Servidor - User:"+lcUserFtp+" - Pass:"+lcPassFtp+" - LOG: " + StrTran(cMsgErrLog,Chr(13)," ")
        fGravaLog( 778, "("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )

        Return(lOk)
    EndIf
EndIf

***********************
** FTP
***********************
If lnPort == 22

    LOCAL lnSuccess, cMsgErrLog, lOk
    LOCAL cResumoLOG, cMensagemLOG

    **//Ligacao SFTP
    loSftp = CreateObject('Chilkat_9_5_0.SFtp')

    **//unlock code do PHC para CHILKAT
    lnSuccess = loSftp.unlockcomponent("PHCPToFTP_LsZMMBcM6rwX")

    **//Inicializa variavel
    cMsgErrLog = "ERRO"
    lOk = .F.

    If Type("loSftp")="O"
        **//Atualiza variavel
        cMsgErrLog = "OK"
        lOk = .T.

        **//Defina alguns tempos limites, em milissegundos:
        loSftp.ConnectTimeoutMs = 5000
        loSftp.IdleTimeoutMs = 10000
    EndIf

    **//Conecte-se ao servidor
    If lOk 
        lnSuccess = loSftp.Connect(lcHostname,lnPort)
    EndIf

    **//Caso ERRO
    If lOk And lnSuccess <> 1
        lOk = .F.
        cMsgErrLog = loSftp.LastErrorText
        RELEASE loSftp

        **//Grava LOG
        cResumoLOG = "ERRO - fLigaSFTP() - Entidade: "+cClientFTP
        cMensagemLOG = "Erro ao conectar no Servidor - Server:"+lcHostname+" - Port:"+aStr(lnPort)+" - LOG: " + StrTran(cMsgErrLog,Chr(13)," ")
        fGravaLog( 778, "("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )

        Return(lOk)
    EndIf

    **// Autentique com o servidor
    If lOk 
        lnSuccess = loSftp.AuthenticatePw(lcUserFtp,lcPassFtp)
    EndIf

    **//Caso ERRO
    If lOk And lnSuccess <> 1
        lOk = .F.
        cMsgErrLog = loSftp.LastErrorText
        RELEASE loSftp

        **//Grava LOG
        cResumoLOG = "ERRO - fLigaSFTP() - Server: "+lcHostname+" - Port:"+aStr(lnPort)
        cMensagemLOG = "Erro ao autenticar no Servidor - User:"+lcUserFtp+" - Pass:"+lcPassFtp+" - LOG: " + StrTran(cMsgErrLog,Chr(13)," ")
        fGravaLog( 778, "("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )

        Return(lOk)
    EndIf

    If lOk
        **// Apos a autenticacao, o subsistema SFTP deve ser inicializado:
        lnSuccess = loSftp.InitializeSftp()
    EndIf

    If lOk And lnSuccess <> 1
        lOk = .F.
        cMsgErrLog = loSftp.LastErrorText
        RELEASE loSftp

        **//Grava LOG
        cResumoLOG = "ERRO - fLigaSFTP() - Server: "+lcHostname+" - Port:"+aStr(lnPort)
        cMensagemLOG = "Erro ao inicializar Servidor - User:"+lcUserFtp+" - Pass:"+lcPassFtp+" - LOG: " + StrTran(cMsgErrLog,Chr(13)," ")
        fGravaLog( 778, "("+cClientFTP+")" ,cResumoLOG ,cMensagemLOG )

        Return(lOk)
    EndIf

EndIf


Return(lOk)
EndFunc

**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fGravaLog
**//|Autor.....: Felipe Aurelio de Melo - melo.felipe@lusocargo.pt
**//|Data......: Junho de 2024
**//|Descricao.: 
**//|Observacao: 
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
    VALUES (	(select left(newid(),25))
                , '<<m_AlertStamp>>' 
                , <<m_AlertId>>
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