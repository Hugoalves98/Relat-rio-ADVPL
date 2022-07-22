#INCLUDE "Protheus.ch"
#INCLUDE "TOPCONN.ch"

#DEFINE ENTER CHR(13)+CHR(10)

user function RELEXCEL()

local oFWMsExcel
local oExcel
local cArquivo    := GetTempPath()+'RelExcel.xml'
local aProdutos   := {}
local aPedidos    := {}
local nX          := 0

MontaQry(@aProdutos,@aPedidos)

oFWMsExcel := FWMsExcel():New()

oFWMsExcel:AddWorkSheet("Produtos")

   //FWMsExcelEx():AddTable(< cWorkSheet >, < cTable >)
   oFWMsExcel:AddTable("Produtos","Produtos")

   //FWMsExcelEx():AddColumn(< cWorkSheet >, < cTable >, < cColumn >, < nAlign >, < nFormat >, < lTotal >)
   oFWMsExcel():AddColumn("Produtos","Produtos","Codigo",1,1)
   oFWMsExcel():AddColumn("Produtos","Produtos","Descricao",1,1)
   oFWMsExcel():AddColumn("Produtos","Produtos","Armazém principal",1,1)

   //FWMsExcelEx():AddRow(< cWorkSheet >, < cTable >, < aRow >,< aCelStyle >)
   for nX := 1 to Len(aProdutos)
        oFWMsExcel():AddRow("Produtos","Produtos",{aProdutos[nX,1],aProdutos[nX,2],aProdutos[nX,3]})
   next

nX := 0
oFWMsExcel:AddWorkSheet("ProdutosVSPedidosCompra")

   //FWMsExcelEx():AddTable(< cWorkSheet >, < cTable >)
   oFWMsExcel:AddTable("ProdutosVSPedidosCompra","Pedidos")

   //FWMsExcelEx():AddColumn(< cWorkSheet >, < cTable >, < cColumn >, < nAlign >, < nFormat >, < lTotal >)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Codigo",1,1)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Descricao",1,1)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Armazém principal",1,1)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Número do pedido de compra",1,1)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Item do pedido",1,1)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Data de emissão",1,4)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Código do fornecedor",1,1)
   oFWMsExcel():AddColumn("ProdutosVSPedidosCompra","Pedidos","Razão Social",1,1)

   //FWMsExcelEx():AddRow(< cWorkSheet >, < cTable >, < aRow >,< aCelStyle >)
   for nX := 1 to Len(aPedidos)
   oFWMsExcel():AddRow("ProdutosVSPedidosCompra","Pedidos",{aPedidos[nX,1],;
                                                            aPedidos[nX,2],;
                                                            aPedidos[nX,3],;
                                                            aPedidos[nX,4],;
                                                            aPedidos[nX,5],;
                                                            Stod(aPedidos[nX,6]),;
                                                            aPedidos[nX,7],;
                                                            aPedidos[nX,8]})
    next
oFWMsExcel:Activate()
oFWMsExcel:GetXMLFile(cArquivo)

oExcel   := MsExcel():New()
oExcel:WorkBooks:Open(cArquivo)
oExcel:SetVisible(.T.)
oExcel:Destroy()

return

static function MontaQry(aProdutos,aPedidos)

local cQuery := ""

cQuery += "SELECT DISTINCT           " + ENTER
cQuery += "B1_COD AS CODIGO,         " + ENTER
cQuery += "B1_DESC AS DESCRICAO,     " + ENTER
cQuery += "B1_LOCPAD AS LOCPAD      " + ENTER
cQuery += "FROM " + retsqlname("SB1") + " B1(NOLOCK)    " + ENTER
cQuery += "INNER JOIN " + retsqlname("SC7") + " C7(NOLOCK) ON B1_COD = C7_PRODUTO AND C7.D_E_L_E_T_ = '' " + ENTER
cQuery += "INNER JOIN " + retsqlname("SA2") + " A2(NOLOCK) ON C7_FORNECE = A2_COD AND C7_LOJA = A2_LOJA AND A2.D_E_L_E_T_ = '' " + ENTER
cQuery += "WHERE B1.D_E_L_E_T_ = '' "

TCQuery cQuery New Alias "QRY1"

while !(QRY1->(Eof()))
    AADD(aProdutos,{QRY1->(CODIGO),QRY1->(DESCRICAO),QRY1->(LOCPAD)})
    QRY1->(DbSkip())
enddo

cQuery := ""

cQuery += "SELECT DISTINCT           " + ENTER
cQuery += "B1_COD AS CODIGO,         " + ENTER
cQuery += "B1_DESC AS DESCRICAO,     " + ENTER
cQuery += "B1_LOCPAD AS LOCPAD,      " + ENTER
cQuery += "C7_NUM AS PEDIDO,         " + ENTER
cQuery += "C7_ITEM AS ITEM,          " + ENTER
cQuery += "C7_EMISSAO AS EMISSAO,    " + ENTER
cQuery += "C7_FORNECE AS FORNECEDOR, " + ENTER
cQuery += "A2_NOME AS NOME          " + ENTER
cQuery += "FROM " + retsqlname("SB1") + " B1(NOLOCK)    " + ENTER
cQuery += "INNER JOIN " + retsqlname("SC7") + " C7(NOLOCK) ON B1_COD = C7_PRODUTO AND C7.D_E_L_E_T_ = '' " + ENTER
cQuery += "INNER JOIN " + retsqlname("SA2") + " A2(NOLOCK) ON C7_FORNECE = A2_COD AND C7_LOJA = A2_LOJA AND A2.D_E_L_E_T_ = '' " + ENTER
cQuery += "WHERE B1.D_E_L_E_T_ = '' "

TCQuery cQuery New Alias "QRY2"

while !(QRY2->(Eof()))
    AADD(aPedidos,{QRY2->(CODIGO),;
                   QRY2->(DESCRICAO),;
                   QRY2->(LOCPAD),;
                   QRY2->(PEDIDO),;
                   QRY2->(ITEM),;
                   QRY2->(EMISSAO),;
                   QRY2->(FORNECEDOR),;
                   QRY2->(NOME)})
    QRY2->(DbSkip())
enddo


return
