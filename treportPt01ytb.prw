#INCLUDE "PROTHEUS.CH"

//DEFINES
#DEFINE ENTER CHR(13)+CHR(10)

//FUNÇÃO PRINCIPAL
user function RELSA1()

Private cPerg := "RELSA1"
Private cNextAlias := GetNextAlias()

ValidPerg(cPerg)

if Pergunte(cPerg,.T.)
   oReport := ReportDef()
   oReport:PrintDialog()
endif

Return 

//SEÇAO DE APRESENTAÇAO DOS DADOS
static function ReportDef()

oReport := TReport():New(cPerg,"Relatório de clientes por Estado",cPerg,{|oReport| ReportPrint(oReport)},"Impressão de relatório de Clientes por Estado")
oReport:SetLandscape(.T.) // Orientação da pag como paisagem
oReport:HideParamPage()

oSection := TRSection():New(oReport,OEMToAnsi("Relatório de Clientes por Estado"),{"SA1"})

//TRCell():New( <oParent> , <cName> , <cAlias> , <cTitle> , <cPicture> , <nSize> , <lPixel> , <bBlock> , <cAlign> , <lLineBreak> , <cHeaderAlign> , <lCellBreak> , <nColSpace> , <lAutoSize> , <nClrBack> , <nClrFore> , <lBold> )
TRCell():New(oSection,"AI_COD",     cNextAlias,"Codigo",    /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 
TRCell():New(oSection,"AI_NOME",    cNextAlias,"Nome",      /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 
TRCell():New(oSection,"PESSOA",     cNextAlias,"Pessoa",    /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 
TRCell():New(oSection,"AI_END",     cNextAlias,"Endereço",  /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 
TRCell():New(oSection,"AI_BAIRRO",  cNextAlias,"Bairro",    /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 
TRCell():New(oSection,"AI_EST",     cNextAlias,"Estado",    /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 
TRCell():New(oSection,"AI_CEP",     cNextAlias,"Cep",       /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 
TRCell():New(oSection,"AI_MUN",     cNextAlias,"Municipio", /*Picture*/,/*Tamanho*/,/*lPixel*/,/*{|| }*/) 


return oReport

//FUNÇAO DE CONSULTA
static function ReportPrint(oReport)

local oSection := oReport:Section(1)
local cQuery   := ""
local nCount   := 0

cQuery += " SELECT    " + ENTER 
cQuery += " A1_COD,  " + ENTER
cQuery += "	A1_NOME, " + ENTER
cQuery += "	CASE WHEN A1_PESSOA = 'J' THEN 'Jurídica' Else 'Física' END PESSOA, " + ENTER
cQuery += "	A1_END, " + ENTER
cQuery += "	A1_BAIRRO, " + ENTER
cQuery += "	A1_EST, " + ENTER
cQuery += "	A1_CEP, " + ENTER
cQuery += "	A1_MUN  " + ENTER
cQuery += "FROM " + RETSQLNAME("SA1") + " WHERE D_E_L_E_T_ = '' " + ENTER 
if !EMPTY(MV_PAR01)
cQuery += " AND A1_EST = '" + MV_PAR01 + "' " + ENTER 
endif
cQuery += "ORDER BY A1_EST, A1_COD "

DbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cNextAlias)

Count to nCount
(cNextAlias)->(dbgotop())
oReport:SetMeter(nCount)
oSection:Init()

while !(cNextAlias)->(Eof())
    oReport:IncMeter()
    oSection:Printline()
    if oReport:cancel()
       exit 
    endif 
    (cNextAlias)->(DbSkip())
enddo

return

//FUNÇAO DE PERGUNTAS
static function ValidPerg(cPerg)

Local aAlias := GetArea()
Local aRegs  := {} 
local i,j 

cPerg := PadR(cPerg,Len(SX1->X1_GRUPO)," ")

AADD(aRegs,{cPerg,"01","Estado","","","mv_ch1", "c",2,0,0,"G","",MV_PAR01,,"","","","","","","","","","","","","","","","","","","","","","","","","12","","","","",""})

DbSelectArea("SX1")
SX1->(DBSetOrder(1))
For i := 1 to LEN(aRegs)
   if !DbSeek(cPerg+aRegs[1,2])
    RecLock("SX1",.T.)
      for j := 1 TO FCount()
      FieldPut(j,aRegs[i,j])
      Next
    MsUnlock()
   ENDIF
next
RestArea(aAlias)

return 


