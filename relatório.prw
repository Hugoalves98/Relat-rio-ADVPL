#INCLUDE 'totvs.ch'
#INCLUDE 'topconn.ch'


User Function RLexcel()

    Processa({||MntQry()},,'Processando...')
    MsAguarde({|| GeraExcel()},,"O arquivo em excel está sendo gerado...")

    DbSelectArea('TR1')
    DBCLOSEAREA()


Return Nil


//função abaixo mostra a query SQL
Static Function MntQry()

    Local cQuery := ""

    
    //pegar dados da base de dados
    cQuery := " SELECT "													
	cQuery += " 	SB1.B1_COD AS CODIGO, "											
	cQuery += " 	SB1.B1_DESC AS DESCRICAO, "										
	cQuery += " 	SB1.B1_TIPO AS TIPO, "										
	cQuery += " 	SBM.BM_GRUPO GRUPO, "										
	cQuery += " 	SBM.BM_DESC BM_DESCRICAO, "										
	cQuery += " 	SBM.BM_PROORI BM_ORIGEM"										
	cQuery += " FROM "													
	cQuery += " 	"+RetSQLName('SB1')+" SB1 "							
	cQuery += " 	INNER JOIN "+RetSQLName('SBM')+" SBM ON ( "		
	cQuery += " 		SBM.BM_FILIAL = '"+FWxFilial('SBM')+"' "		
	cQuery += " 		AND SBM.BM_GRUPO = B1_GRUPO "					
	cQuery += " 		AND SBM.D_E_L_E_T_='' "							
	cQuery += " 	) "														
	cQuery += " WHERE "													
	cQuery += " 	SB1.B1_FILIAL = '"+FWxFilial('SBM')+"' "			
	cQuery += " 	AND SB1.D_E_L_E_T_ = '' "							
	cQuery += " ORDER BY "												
	cQuery += " 	SB1.B1_COD "

        If select('TR1') <> 0
            DbSelectArea('TR1')
            DBCLOSEAREA(  )
        ENDIF

        cQuery := ChangeQuery(cQuery)
        DBUSEAREA( .T.,'TOPCONN',TCGENQRY(,,cQuery),'TR1',.F.,.T.)

Return Nil



Static Function GeraExcel()
    
    Local oExcel := FWMSEXCEL():New()
    Local lOK := .F.
    Local cArq := ""
    Local cDirTMP := "C:\spool\"

    DbSelectArea('TR1')
    TR1->(DBGOTOP( ))

    oExcel:AddWorkSheet("PRODUTOS")
	oExcel:AddTable("PRODUTOS","TESTE")
	oExcel:AddColumn("PRODUTOS","TESTE","CODIGO",1,1)
	oExcel:AddColumn("PRODUTOS","TESTE","DESCRICAO",1,1)
	oExcel:AddColumn("PRODUTOS","TESTE","TIPO",1,1)
	oExcel:AddColumn("PRODUTOS","TESTE","GRUPO",1,1)
	oExcel:AddColumn("PRODUTOS","TESTE","BM_DESCRICAO",1,1)
	oExcel:AddColumn("PRODUTOS","TESTE","BM_ORIGEM",1,1)
	


        While TR1->(!EOF())

            oExcel:AddRow("Produtos","teste",{TR1->(CODIGO),;
                                                TR1->(DESCRICAO),;
                                                TR1->(TIPO),;
                                                TR1->(GRUPO),;
                                                TR1->(BM_DESCRICAO),;
                                                TR1->(BM_ORIGEM)})
            lOK:=.T.
            TR1->(dbskip())

        ENDDO
    oExcel:Activate()

        cArq := CriaTrab(NIL, .F.) + '.xml'
        oExcel:GetXMLFile(cArq)

            If __Copyfile( cArq, cDirTMP+  cArq)
                if lOk 
                    oExcelApp := MSExcel():New()
                    oExcelApp:WorkBooks:Open(cDirTMP+ cArq)
                    oExcelApp:SetVisible(.T.)
                    oExcelApp:Destroy()
                
                MsgIndfo('O arquivo Excel foi gerado no diretório:'+cDirTMP+ cArq+". ")


                ENDIF
                else
                        MsgAlert('Deu ruim ao copiar o arquivo excel', 'teste')

            ENDIF

Return
