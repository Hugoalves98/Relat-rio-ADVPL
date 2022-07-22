#INCLUDE "Protheus.ch"
#INCLUDE "parmtype.ch"
#INCLUDE "TOPCONN.ch"

user function geraRel()

Private cCliente1 := Space(6)
Private cCliente2 := Space(6)

getParam()



return



static function getParam()

local alParamBox      := {}
local clTitulo     := "Parâmetros"
local alButtons  := {}
local llCentered := .T. 
Local nlPosx  := Nil
Local nlPosy := Nil
Local clLoad := ""
Local llCanSave   := .T.
Local llUserSave := .T.
Local llRet := .T.
Local blOk
Local alParams  := {}

AADD(alParamBox,{1,"Cliente De"     ,space(6)      ,"@!",".T."     ,"SA1"    ,"",25,.F.})
AADD(alParamBox,{1,"Cliente até"     ,space(6)      ,"@!",".T."     ,"SA1"    ,"",25,.F.})

llRet := ParamBox(alParamBox, clTitulo, alParams, blOk, alButtons, llCentered, nlPosx, nlPosy,, clLoad, llCanSave, llUserSave)


return


