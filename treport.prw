//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"
	
//Constantes
#Define STR_PULA		Chr(13)+Chr(10)
	
/*/{Protheus.doc} TRTAPX02
Relatório - Planejamento de Remessa       
@author zReport
@since 27/06/2022
@version 1.0
	@example
	u_TRTAPX02()
	@obs Função gerada pelo zReport()
/*/
	
User Function TRTAPX02()
	Local aArea   := GetArea()
	Local oReport := NIL
	Local lEmail  := .F.
	Local cPara   := ""
	Private cPerg := Padr('TRPTAPX01',10)
    Pergunte(cPerg,.F.) //SX1
	
	//Cria as definições do relatório
	oReport := fReportDef()
	
	//Será enviado por e-Mail?
	If lEmail
		oReport:nRemoteType := NO_REMOTE
		oReport:cEmail := cPara
		oReport:nDevice := 3 //1-Arquivo,2-Impressora,3-email,4-Planilha e 5-Html
		oReport:SetPreview(.F.)
		oReport:Print(.F., "", .T.)
	//Senão, mostra a tela
	Else
		oReport:PrintDialog()
	EndIf
	
	RestArea(aArea)
Return
	
/*-------------------------------------------------------------------------------*
 | Func:  fReportDef                                                             |
 | Desc:  Função que monta a definição do relatório                              |
 *-------------------------------------------------------------------------------*/
	
Static Function fReportDef()
	Local oReport
	Local oSectDad := Nil
	Local oSectDadLog := Nil
	Local oBreak := Nil
	Local oBreak1 := Nil
    Local simbol := '>'
		
	//Criação do componente de impressão
	oReport := TReport():New(	"TRTAPX02",;		//Nome do Relatório
								"Planejamento de Remessa",;		//Título
								cPerg,;		//Pergunte ... Se eu defino a pergunta aqui, será impresso uma página com os parâmetros, conforme privilégio 101
								{|oReport| fRepPrint(oReport)},;		//Bloco de código que será executado na confirmação da impressão
								)		//Descrição
	oReport:SetTotalInLine(.F.)
	oReport:lParamPage := .F.
	oReport:oPage:SetPaperSize(9) //Folha A4
	oReport:SetLandscape()
		
	//Criando a seção de dados
	oSectDad := TRSection():New(	oReport,;		//Objeto TReport que a seção pertence
									"Dados",;		//Descrição da seção
									{"QRY_AUX"})		//Tabelas utilizadas, a primeira será considerada como principal da seção
	oSectDad:SetTotalInLine(.F.)  //Define se os totalizadores serão impressos em linha ou coluna. .F.=Coluna; .T.=Linha
	
	//Colunas da primeira seção
	TRCell():New(oSectDad, "CODIGO", "QRY_AUX", "Codigo", /*Picture*/, 10, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "MES", "QRY_AUX", "Data", /*Picture*/, 8, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "TIPO_REMESSA", "QRY_AUX", "Tipo", /*Picture*/, 10, /*lPixel*/,{|| IIF(("QRY_AUX")->TIPO_REMESSA== "1","Gestao","Promocao")},/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.F.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "CENTRO_CUSTO", "QRY_AUX", "Centro de custo", /*Picture*/, 9, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "DESC_CC", "QRY_AUX", "Descricao CC", /*Picture*/, 30, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad, "INICIATIVA", "QRY_AUX", "Iniciativa", /*Picture*/, 6, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad, "DESCRICAO", "QRY_AUX", "Descricao", /*Picture*/, 30, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "TIPO_MOEDA", "QRY_AUX", "Tipo_moeda", /*Picture*/, 1, /*lPixel*/,{|| IIF(("QRY_AUX")->TIPO_MOEDA== "1","Euro","Dolar")},/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad, "VAL_INI", "QRY_AUX", "Val_ini", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,"RIGHT",/*lLineBreak*/,"RIGHT",/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "SALD_INI", "QRY_AUX", "Sald_ini", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,"RIGHT",/*lLineBreak*/,"RIGHT",/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "VAL_PLANREM", "QRY_AUX", "Val_planrem", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,"RIGHT",/*lLineBreak*/,"RIGHT",/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "VAL_PLANREMAT", "QRY_AUX", "Val_planremat", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,"RIGHT",/*lLineBreak*/,"RIGHT",/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "VAL_EXEC", "QRY_AUX", "Val_exec", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,"RIGHT",/*lLineBreak*/,"RIGHT",/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad, "SALDO", "QRY_AUX", "Saldo", /*Picture*/, 15, /*lPixel*/,/*{|| code-block de impressao }*/,"RIGHT",/*lLineBreak*/,"RIGHT",/*lCellBreak*/,/*nColSpace*/,.T.,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	
	oSectDadLog :=TRSection():New( oReport,;
									"Logs e Alterações",;
									{"QRY_AUX"})

	oSectDadLog:SetTotalInLine(.F.) 
	TRCell():New(oSectDadLog, "LOG_INCLU", "QRY_AUX", "Log de inclusao", /*Picture*/, 10, /*lPixel*/,{||FWLEUSERLG("LOG_INCLU")},/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDadLog, "LOG_ALTER", "QRY_AUX", "Log de alteracao", /*Picture*/, 8, /*lPixel*/,{||FWLEUSERLG("LOG_ALTER")},/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	
	oBreak:=TRBreak():New( oSectDad ,{|| QRY_AUX->(CODIGO)} , {||"Dados"}) 
	oSectDad:SetHeaderBreak(.T.)

	oBreak1:=TRBreak():New( oSectDadLog ,{|| QRY_AUX->(CODIGO)},{||"Secao de Logs e alteracoes"}) 
	oSectDad:SetTitle("Dados")
	oSectDadLog:SetHeaderBreak(.T.)

	oSectDadLog:SetHeaderSection(.T.)
	oSectDad:SetPageBreak(.T.)
Return oReport
	
/*-------------------------------------------------------------------------------*
 | Func:  fRepPrint                                                              |
 | Desc:  Função que imprime o relatório                                         |
 *-------------------------------------------------------------------------------*/
	
Static Function fRepPrint(oReport)
	Local aArea    := GetArea()
	Local cQryAux  := ""
	Local oSectDad := Nil
	Local oSectDadLog := Nil
	Local nAtual   := 0
	Local nTotal   := 0
    Local cCombo := ""
	Local cNumCod:=""
	Local ca:={||Posicione("QRY_AUX",1,XFIlial("QRY_AUX")+M->CODIGO,FWLEUSERLG("LOG_INCLU"))}

	//Pegando as seções do relatório
	oSectDad := oReport:Section(1)
	oSectDadLog := oReport:Section(2)

    //Condicional para validar filtro do saldo
    if MV_PAR11 = 1
        cCombo:="="
    ELSEIF MV_PAR11 = 2
        cCombo:=">"
    elseif MV_PAR11 = 3
        cCombo:=">="
    endif

	//Montando consulta de dados
	cQryAux := ""
	cQryAux += "SELECT ZBC_CODIGO AS CODIGO,"		+ STR_PULA
	cQryAux += "       ZBC_MESREM AS MES,"		+ STR_PULA
	cQryAux += "       ZBC_TIPREM AS TIPO_REMESSA,"		+ STR_PULA
	cQryAux += "       ZBC_CCUSTO AS CENTRO_CUSTO,"		+ STR_PULA
	cQryAux += "       CTT_XEAD   AS DESC_CC,"		+ STR_PULA
	cQryAux += "       ZBC_TPMOED AS TIPO_MOEDA,"		+ STR_PULA
	cQryAux += "       ZBC_INICIA AS INICIATIVA,"		+ STR_PULA
	cQryAux += "       ZA_DESC    AS DESCRICAO,"		+ STR_PULA
	cQryAux += "       ZBC_VALINI AS VAL_INI,"		+ STR_PULA
	cQryAux += "       ZBC_SALINI AS SALD_INI,"		+ STR_PULA
	cQryAux += "       ZBC_VALREM AS VAL_PLANREM,"		+ STR_PULA
	cQryAux += "       ZBC_VALREA AS VAL_PLANREMAT,"		+ STR_PULA
	cQryAux += "       ZBC_VALEXE AS VAL_EXEC,"		+ STR_PULA
	if MV_PAR12 == 1
		cQryAux += "       ZBC_SALDO  AS SALDO,"		+ STR_PULA
		cQryAux += "       ZBC_USERGI AS LOG_INCLU,"		+ STR_PULA
		cQryAux += "       ZBC_USERGA AS LOG_ALTER"		+ STR_PULA
	ELSE
	cQryAux += "       ZBC_SALDO  AS SALDO"		+ STR_PULA
	ENDIF
	cQryAux += "	   FROM  ZBC990 AS ZBC"		+ STR_PULA
	cQryAux += "       INNER JOIN SZA990 AS ZA"		+ STR_PULA
	cQryAux += "               ON ZBC.ZBC_VALINI = ZA_VALINI "		+ STR_PULA
	cQryAux += "                  AND ZBC_SALINI = ZA.ZA_SALINI "		+ STR_PULA
	cQryAux += "                  AND ZA.D_E_L_E_T_ = ''"		+ STR_PULA
	cQryAux += "       INNER JOIN CTT990 AS CT"		+ STR_PULA
	cQryAux += "               ON ZBC.ZBC_CCUSTO = CT.CTT_CUSTO"+ STR_PULA
	cQryAux += "                  AND CT.D_E_L_E_T_ = ''"		+ STR_PULA
    cQryAux += "                  WHERE  ZBC.D_E_L_E_T_ = ''"		+ STR_PULA
	cQryAux += "                    AND ZBC_CODIGO BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"+STR_PULA
	cQryAux += "                    AND ZBC_MESREM  BETWEEN '"+DTOS(MV_PAR03)+"' AND '"+DTOS(MV_PAR04)+"'"+STR_PULA
	cQryAux += "                    AND ZBC_CCUSTO BETWEEN'"+MV_PAR06+"' AND '"+MV_PAR07+"'"+STR_PULA
	cQryAux += "                    AND ZBC_INICIA BETWEEN'"+MV_PAR09+"'AND'"+MV_PAR10+"'"+STR_PULA
    cQryAux += "                    AND ZBC_SALDO "+cCombo+" 0" +STR_PULA
    if !(MV_PAR05 == 3) .and. !(MV_PAR08 == 3)    
        cQryAux += "                    AND ZBC_TIPREM = '"+cvaltochar(MV_PAR05)+"'"+STR_PULA
        cQryAux += "                    AND ZBC_TPMOED = '"+cvaltochar(MV_PAR08)+"'"+STR_PULA
    ELSEIF MV_PAR05 == 3 .AND. !(MV_PAR08 == 3)
        cQryAux += "                    AND ZBC_TIPREM = '1' OR ZBC_TIPREM = '2'"+STR_PULA
        cQryAux += "                    AND ZBC_TPMOED = '"+cvaltochar(MV_PAR08)+"'"+STR_PULA
    ELSEIF !(MV_PAR05 == 3) .AND. MV_PAR08 == 3
        cQryAux += "                    AND ZBC_TIPREM = '"+cvaltochar(MV_PAR05)+"'"+STR_PULA
        cQryAux += "                    AND ZBC_TPMOED = '2' OR ZBC_TPMOED = '1'"+STR_PULA
     Elseif MV_PAR05 == 3 .AND. MV_PAR08 == 3
        cQryAux += "                    AND ZBC_TIPREM = '1' OR ZBC_TIPREM = '2'"+STR_PULA
        cQryAux += "                    AND ZBC_TPMOED = '2' OR ZBC_TPMOED = '1'"+STR_PULA
    ENDIF
    cQryAux += "ORDER  BY CODIGO,"		+ STR_PULA
	cQryAux += "          MES,"		+ STR_PULA
	cQryAux += "          TIPO_REMESSA,"		+ STR_PULA
	cQryAux += "          CENTRO_CUSTO,"		+ STR_PULA
	cQryAux += "          DESC_CC,"		+ STR_PULA
	cQryAux += "          INICIATIVA,"		+ STR_PULA
	cQryAux += "          DESCRICAO,"		+ STR_PULA
	cQryAux += "          TIPO_MOEDA,"		+ STR_PULA
	cQryAux += "          VAL_INI,"		+ STR_PULA
	cQryAux += "          SALD_INI,"		+ STR_PULA
	cQryAux += "          VAL_PLANREM,"		+ STR_PULA
	cQryAux += "          VAL_PLANREMAT,"		+ STR_PULA
	cQryAux += "          VAL_EXEC,"		+ STR_PULA
	cQryAux += "          SALDO,"		+ STR_PULA
	IF MV_PAR12 ==1
		cQryAux += "          LOG_INCLU,"		+ STR_PULA
		cQryAux += "          LOG_ALTER;"		+ STR_PULA
	
	ELSE
	cQryAux += "          SALDO;"		+ STR_PULA	
	ENDIF
	cQryAux := ChangeQuery(cQryAux)
	

	//Executando consulta e setando o total da régua
	TCQuery cQryAux New Alias "QRY_AUX"
	
	Count to nTotal
	oReport:SetMeter(nTotal)
	TCSetField("QRY_AUX", "MES", "D")
	TCSetField("QRY_AUX","VAL_INI","N",14)
	
	//Enquanto houver dados
	oSectDad:Init()
	QRY_AUX->(DbGoTop())
	//MsgAlert(VALTYPE(QRY_AUX->CODIGO))
	//MsgAlert(CValToChar(ca))
	//MsgAlert(FWLEUSERLG('LOG_INCLU'))
	While ! QRY_AUX->(Eof())
		//Incrementando a régua
		nAtual++
		cNumCod:= QRY_AUX->CODIGO
		oReport:SetMsgPrint("Imprimindo registro "+cValToChar(nAtual)+" de "+cValToChar(nTotal)+"...")
		oReport:IncMeter()
		
		//Imprimindo a linha atual
		oSectDad:PrintLine()

		oSectDadLog:Init()
		While QRY_AUX->CODIGO==cNumCod
			
		
		IncProc("Imprimindo Logs...."+ AllTrim(QRY_AUX->LOG_INCLU))
		oReport:IncMeter()
		oSectDadLog:Printline()
		QRY_AUX->(DbSkip())
		ENDDO
		
	EndDo
	oSectDadLog:Finish()
	oSectDad:Finish()
	QRY_AUX->(DbCloseArea())
	
	RestArea(aArea)
Return
