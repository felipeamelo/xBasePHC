**//+-----------------------------------------------------------------------------------//
**//|Funcao....: fGeraPDFeXMLfatura(cFTstamp)
**//|Autor.....: Felipe Aurélio de Melo
**//|Data......: 14 de março de 2025, 09:00
**//|Descricao.: Gera o PDF da Fatura e o XML com PDF dentro do XML
**//|Observacao: Basta chamar a função passando o Stamp do FT da Fatura em questão
**//|          : e a função retorna o diretorio e nome do ficheiro PDF gerado
**//|          : Ficheiros são gerando em: C:\PHC\PDF\FaturasPDF\*.*
**//+-----------------------------------------------------------------------------------//
**-----------------------------------------------------------*
** Function fGeraPDFeXMLfatura(cFTstamp)
**-----------------------------------------------------------*
Lparameters cParamFTstamp

** Declaração de variaveis
Local cTipoImpress,cStampFatura,nNumSerie,cNomeSerie,nNumFatura,nAnoFatura,cDirPDF,cArqNew, c_nomexml, c_nomefilesign, tcfilepdf
Store "" To cTipoImpress, cStampFatura, cNomeSerie, cDirPDF, cArqNew, cFileDel, m.c_nomexml, m.c_nomefilesign, m.tcfilepdf
Store 0 To nNumSerie, nNumFatura, nAnoFatura

** Restaura Diretorio Padrão
Set Default To alltrim( JUSTPATH ( MAIN_DIR ) )

** Inicializa variaveis
cDirPDF = Alltrim(SYS(5)) + AddBS(CurDir()) + AddBS("PDF\FaturasPDF\")

** Cria diretorio caso não exista
If !Directory(cDirPDF)
	MkDir(cDirPDF)
EndIf

** Verifica se tabela esta preparada para utilizar
if not used("FT")
	do dbfuseFT
endif
if not used("FT2")
	do dbfuseFT2
endif
if not used("FT3")
	do dbfuseFT3
endif
if not used("FI")
	do dbfuseFI
endif
if not used("FI2")
	do dbfuseFI2
endif
If Not Used("FTT")
	do dbfuseFTT
Endif

Do dbfusecidu With "FTIDUC" 
Do dbfuselidu With "FTIDUC", "FTIDUL" 

** Seleciona Tabela a ser utilizada
Select FT

** Define Stamp a ser impresso
cStampFatura = Alltrim(cParamFTstamp) &&Alltrim(FT.ftstamp)

** Carregar Views
v_ftstamp = cStampFatura
u_requery("FT")

v_ft2ftstamp = cStampFatura
u_requery("FT2")

v_ft3ftstamp = cStampFatura
u_requery("FT3")

v_fttstamp = cStampFatura
u_requery("FTT")

v_fiftstamp = cStampFatura
u_requery("FI")

v_fi2ftstamp = cStampFatura
u_requery("FI2")

** Define nome do ficheiro PDF (Serie +  Fatura + Ano)
Select FT
nNumSerie    = FT.ndoc
cNomeSerie   = Alltrim(FT.nmdoc)
nNumFatura   = FT.fno
nAnoFatura   = FT.ftano
cArqNew      = aStr(nNumSerie)+aStr(nNumFatura)+aStr(nAnoFatura)

** Define o tipo de impressão
Do Case
	Case Alltrim(FT.u_via) == 'R'
		cTipoImpress = "Rodoviária - papel branco"

	Case Alltrim(FT.u_via) == 'D'
		cTipoImpress = "Nacional - papel branco"

	Case Alltrim(FT.u_via) == 'M'
		cTipoImpress = "Maritima - papel branco"

	Case Alltrim(FT.u_via) == 'A'
		cTipoImpress = "Aérea - papel branco"

	Case Alltrim(FT.u_via) == 'O'
		cTipoImpress = "Outros - papel branco"

	Case Alltrim(FT.u_via) == 'H'
		cTipoImpress = "Despachos - papel branco"

	Case Alltrim(FT.u_via) == 'E'
		cTipoImpress = "Entreposto - papel branco"

	Case Alltrim(FT.u_via) == 'P'
		cTipoImpress = "Aduaneira - papel branco"

	Case Alltrim(FT.u_via) == 'L'
		cTipoImpress = "Armazém 3 - papel branco"
EndCase

** Posiciona no IDU que será utilizado para impressão
Select ftlist
m.v_idundos = nNumSerie
Select FTIDUC
u_REQUERY("FTIDUC")
Locate For FTIDUC.titulo=cTipoImpress

v_iducstamp=FTIDUC.idustamp
v_idulstamp=FTIDUC.idustamp
Select FTIDUL
u_REQUERY("FTIDUL")

Do tdread With "", nNumSerie

** chama Impressão do IDU
report_dir=cDirPDF
IduToPdf("FT","FI","FTCAMPOS","FICAMPOS","FTIDUC","FTIDUL",nNumSerie,cTipoImpress,cDirPDF+cArqNew+".pdf","","NO",.F.,"ONETOMANY","",1,3)

** Apaga ficheiro gerado automaticamente
cFileDel = cDirPDF+cNomeSerie+" nº "+aStr(nNumFatura)+".pdf"
If File(cFileDel)
	InKey(2)
	Delete File (cFileDel)
Endif

** Gera XML com PDF dentro
m.tcfilepdf = cDirPDF+cArqNew+".pdf"
makexmlubl_2_1_cius(@c_nomexml, @c_nomefilesign, tcfilepdf, y_tiposaft)
Rename (c_nomexml) To (cDirPDF+cArqNew+".xml"))

** Limpar Views
v_ftstamp = "-+-"
u_requery("FT")

v_ft2ftstamp = "-+-"
u_requery("FT2")

v_ft3ftstamp = "-+-"
u_requery("FT3")

v_fttstamp = "-+-"
u_requery("FTT")

v_fiftstamp = "-+-"
u_requery("FI")

v_fi2ftstamp = "-+-"
u_requery("FI2")

** Restaura Diretorio Padrão
Set Default To alltrim( JUSTPATH ( MAIN_DIR ) )

Return(cDirPDF+cArqNew+"pdf")
