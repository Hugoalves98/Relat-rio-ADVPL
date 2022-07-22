User Function Re1Txt()
    If MsgYesNo('Esta função tem como objetivo gerar arquivo txt')
        Processa({ MntQry()},,'Processando...')
        MsAguarde({ GeraArqi()},,'O arquivo TXT está sendo gerado....')
    ELSE
        Alert('Cancelada pelo operador')

    ENDIF

Return Nil


Static function GeraArqi()

    Local nHandle := FCreate( 'C:/Users/Yuri/Desktop/Treinamento/relatorios/arquivo2.txt' )
    //Local nlinha 

    if nHandle <0
        MsgInfo('Erro ao criar o arquivo', 'ERRO')
    else
        //for nlinha :=1 to 200
            //Fwrite(nHandle,"Gravando a linha" + Strzero(nlinha,3)+ CRLF)
       //NEXT nlinha
       While TMP->(!EOF())
            FWRITE( nHandle, TMP->(CODIGO) + '|' + TMP->(FANTASIA) +'|'+ TMP->(RAZAO) + CRLF )
            TMP->(dbskip())

       ENDDO


        Fclose(nHandle)
    ENDIF

    if FILE('C:/Users/Yuri/Desktop/Treinamento/relatorios/arquivo2.txt')
        MsgInfo('Arquivo criado com sucesso')

    Else
        MsgAlert('N foi possivel criar o arquivo', 'Alerta')
    ENDIF


Return


Static Function MntQry()
    Local cQuery := ''
    cQuery := 'SELECT ZZ1_CODIGO AS CODIGO, '
    cQuery += 'ZZ1_FANTAS AS FANTASIA, ZZ1_RAZAOS AS RAZAO'
    cQuery += "FROM ZZ1990 WHERE D_E_L_ET = '' "

    cQuery := ChangeQuery(cQuery)
        DBUSEAREA( .T.,'TOPCONN', TCGENQRY(,,cQuery), 'TMP',.F.,.T.)

Return Nil
