Attribute VB_Name = "modEscrituracao"
Option Compare Database

Public Conn As New ADODB.Connection
Public SQLStr As String


Public Sub ConnectToDataBase()
 
 Dim Server_Name As String
 Dim Database_Name As String
 Dim User_ID As String
 Dim Password As String
 '
 Dim I As Long ' counter
 ' SQL to perform various actions
 Dim table1 As String, table2 As String
 Dim field1 As String, field2 As String
 Dim rs As ADODB.Recordset
 Dim vtype As Variant
 
 
Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()
Set rst = Db.OpenRecordset("tbDBConn")



 Server_Name = rst!host ' Enter your server name here -
 Database_Name = rst!dbname ' Enter your database name here
 User_ID = rst!User ' enter your user ID here
 Password = rst!pass ' Enter your password here

 Set Conn = New ADODB.Connection
Conn.Open "DRIVER={MySQL ODBC 8.0 Unicode Driver}" _
 & ";SERVER=" & Server_Name _
 & ";DATABASE=" & Database_Name _
 & ";UID=" & User_ID _
 & ";PWD=" & Password _
 & ";OPTION=16427" ' Option 16427 = Convert LongLong to Int:
End Sub

Sub DisconnectFromDataBase()

Close Connections
On Error Resume Next
rs.Close
Set rs = Nothing
Conn.Close
Set Conn = Nothing
On Error GoTo 0

End Sub
Public Sub Calcula_Custo_Medio(cDtIni As String, cDtFim As String)

Call ConnectToDataBase

Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

Dim cSTR_DtINI As String
Dim cSTR_DtFIM As String

cSTR_DtINI = Replace(Format(cDtIni, "dd/mm/yyyy"), "/", "")
cSTR_DtFIM = Replace(Format(cDtFim, "dd/mm/yyyy"), "/", "")


cDtIni = Format(cDtIni, "yyyy-mm-dd")
cDtFim = Format(cDtFim, "yyyy-mm-dd")

cDtIniVb = Format(cDtIni, "mm/dd/yyyy")
cDtFimVb = Format(cDtFim, "mm/dd/yyyy")



strSQL = ("UPDATE tbVendasDet SET tbVendasDet.VlrLiquido = tbVendasDet.ValorTot-tbVendasDet.Valor_ICMS-tbVendasDet.Valor_PIS-tbVendasDet.Valor_Cofins-tbVendasDet.Valor_IPI-tbVendasDet.Valor_ICMS_ST " & _
"WHERE tbVendasDet.VlrLiquido=0;")
Conn.Execute strSQL

strSQL = ("update tbCadprod set Litros = 1 where Litros is null;")
Conn.Execute strSQL
'CALCULA CUSTO DE PRODUCAO

'CALCULA CUSTO DE PRODUCAO
strSQL = ("update tbcadprod set CMed_Unit = 0, CMed_Tot = 0 where CMed_Unit = null;")
Conn.Execute strSQL


'BUSCA O CUSTO UNITARIO DE PRODUTOS DE REVENDA COM TROCA DE CÓDIGO E GRANEL
Set rst = Db.OpenRecordset("SELECT tbVendas.ANO, tbVendas.MES, tbVendas.DataEmissao, tbVendasDet.IDProd, tbVendasDet.Qnt, tbVendasDet.CustoMedio, tbVendasDet.VlrLiquido, tbCadProd.CMed_Unit, tbCadProd.DescProd, tbCadProd.LITROS, tbCadProd.IdProd_Revenda, tbcadprod_1.DescProd,tbcadprod_1.Litros as LitrosOriginal, tbcadprod_1.CMed_Unit as CMed_Unit_Original, tbcadprod_1.Estoque " & _
"FROM (tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) LEFT JOIN tbcadprod AS tbcadprod_1 ON tbCadProd.IdProd_Revenda = tbcadprod_1.IDProd " & _
"WHERE tbVendas.DataEmissao >= #" & cDtIniVb & "# And tbVendas.DataEmissao <= #" & cDtFimVb & " 23:59:59" & "# AND tbVendasDet.CFOP<>'1604';")

'"WHERE (((tbVendas.DataEmissao) >= #" & cDtIniVB & "# And (tbVendas.DataEmissao) <= #" & cDtfIMVb & "#)) " & _

Do Until rst.EOF
'On Error Resume Next
If IsNull(rst.Fields("CMed_Unit_Original").Value) Then
'IsNull (rs.Fields("MiddleInitial").Value)

Else
strSQL = ("UPDATE tbCadProd SET tbCadProd.Estoque = 0, tbCadProd.CMed_Unit = " & Replace((rst!CMed_Unit_Original / rst!LitrosOriginal) * rst!Litros, ",", ".") & " where tbCadProd.IDProd = " & rst!IDProd & ";")
Conn.Execute strSQL
End If


rst.MoveNext
Loop
'BUSCA O CUSTO UNITARIO DE PRODUTOS DE REVENDA COM TROCA DE CÓDIGO E GRANEL
On Error GoTo 0

'REGISTRA CUSTO MEDIO DA VENDA E VALOR LIQUIDO
strSQL = ("UPDATE tbVendasDet SET tbVendasDet.VlrLiquido = tbVendasDet.ValorTot-tbVendasDet.Valor_ICMS-tbVendasDet.Valor_PIS-tbVendasDet.Valor_Cofins-tbVendasDet.Valor_IPI-tbVendasDet.Valor_ICMS_ST " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE tbVendas.DataEmissao >= '" & cDtIni & "' And tbVendas.DataEmissao <= '" & cDtFim & " 23:59:59" & "';")
'Conn.Execute strSQL

strSQL = ("update tbVendasDet as q1 inner join tbCadProd as q2 on q1.IdProd = q2.IdProd Set q1.CustoMedio = q2.CMed_Unit * q1.Qnt where CustoMedio = 0")
Conn.Execute strSQL

'REGISTRA CUSTO MEDIO DA VENDA E VALOR LIQUIDO




'CORRIGE CUSTO DE PRODUÇÃO DE PRODUTOS PRODUZIDOS
Call Calc_CustoProducao
'CORRIGE CUSTO DE PRODUCAO DE PRODUTOS PRODUZIDOS
Call ConnectToDataBase

strSQL = ("UPDATE tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd SET tbVendasDet.CustoMedio = tbCadProd.CMed_Unit*tbVendasDet.Qnt " & _
"FROM (tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) LEFT JOIN tbcadprod AS tbcadprod_1 ON tbCadProd.IdProd_Revenda = tbcadprod_1.IDProd " & _
"WHERE tbVendas.DataEmissao >= '" & cDtIni & "' And tbVendas.DataEmissao <= '" & cDtFim & " 23:59:59" & "';")
'Conn.Execute strSQL

strSQL = ("UPDATE tbVendasDet as q1 " & _
"inner join tbVendas as q2 on q2.ID = q1.IDVenda " & _
"Set Margem = Round(((q1.VlrLiquido - q1.CustoMedio) / q1.VlrLiquido) * 100, 2) " & _
"WHERE DataEmissao >= '" & cDtIni & "' And DataEmissao <= '" & cDtFim & " 23:59:59" & "';")
Conn.Execute strSQL



'ajustes
'REGISTRA SALDOS DE ESTOQUE
'strSQL = ("UPDATE tbCadProd set Estoque = 0, CMed_Unit = 0, CMed_Tot=0 ")
'Conn.Execute strSQL
'entradas +
Set rstSaldo = Db.OpenRecordset("SELECT tbComprasDet.IDProd, Sum(tbComprasDet.Qnt) AS QT, Sum(tbComprasDet.ValorTot) AS ValorTot FROM tbComprasDet GROUP BY tbComprasDet.IDProd;")
    Do Until rstSaldo.EOF
    strSQL = ("UPDATE tbCadProd set Estoque = " & Replace(rstSaldo!Qt, ",", ".") & ", CMed_Unit = " & Replace(rstSaldo!ValorTot / rstSaldo!Qt, ",", ".") & ", CMed_Tot = " & Replace(rstSaldo!ValorTot, ",", ".") & " where IdProd = " & rstSaldo!IDProd & "")
    Conn.Execute strSQL
    rstSaldo.MoveNext
    Loop
'saidas Vendas -

Set rstSaldo = Db.OpenRecordset("SELECT tbVendasDet.IDProd, Sum(tbVendasDet.Qnt) AS QT, Sum(tbVendasDet.CustoMedio) AS ValorTot FROM tbVendasDet where CFOP <> '1551' and CFOP <> '5920' GROUP BY tbVendasDet.IDProd ; ")
    Do Until rstSaldo.EOF
'    strSQL = ("UPDATE tbCadProd set Estoque = Estoque - " & Replace(rstSaldo!Qt, ",", ".") & ", CMed_Tot = CMed_Tot - " & Replace(rstSaldo!ValorTot, ",", ".") & ", CMed_Unit = (CMed_Tot - " & Replace(rstSaldo!ValorTot, ",", ".") & " ) / (Estoque - " & Replace(rstSaldo!Qt, ",", ".") & ") where IdProd = " & rstSaldo!IDProd & "")
 '   Conn.Execute strSQL
    rstSaldo.MoveNext
    Loop
    ''DoCmd.setwarnings (False)
'ajustes manuais + ou -
'consumos e perdas diversas -
'producao +
Set rstSaldo = Db.OpenRecordset("SELECT tb_Ajuste_Estoque.IDProd, tb_Ajuste_Estoque.QT, tb_Ajuste_Estoque.CMed_Tot, tb_Ajuste_Estoque.CMed_Unit FROM tb_Ajuste_Estoque order by Data Asc;")
    Do Until rstSaldo.EOF
    strSQL = ("UPDATE tbCadProd set Estoque = Estoque + " & Replace(rstSaldo!Qt, ",", ".") & ", CMed_Tot = CMed_Tot + " & Replace(rstSaldo!CMed_Tot, ",", ".") & " where IdProd = " & rstSaldo!IDProd & "")
    Conn.Execute strSQL
    ', CMed_Unit = (CMed_Tot + " & Replace(rstSaldo!ValorTot, ",", ".") & " ) / (Estoque + " & Replace(rstSaldo!QT, ",", ".") & ")
    strSQL = ("UPDATE tbCadProd set CMed_Unit = CMed_Tot / Estoque where IdProd = " & rstSaldo!IDProd & " and Estoque > 0 ")
    Conn.Execute strSQL
    strSQL = ("UPDATE tbCadProd set CMed_Unit = " & Replace(rstSaldo!CMed_Unit * -1, ",", ".") & " where IdProd = " & rstSaldo!IDProd & " and Estoque < 0 ")
    Conn.Execute strSQL
    strSQL = ("UPDATE tbCadProd set CMed_Unit = " & Replace(rstSaldo!CMed_Unit * -1, ",", ".") & ", CMed_Tot = 0 where IdProd = " & rstSaldo!IDProd & " and Estoque = 0 ")
    Conn.Execute strSQL
    
    rstSaldo.MoveNext
    Loop

'corrigir negativos unitarios estoque zero
strSQL = ("update tbCadProd set CMed_Unit = CMed_Unit * -1 where Estoque = 0 and CMed_Unit < 0")
Conn.Execute strSQL

'zerar saldos de energia
strSQL = ("UPDATE tbCadProd set Estoque = 0, CMed_Tot=0 where idProd = 2800")
Conn.Execute strSQL
'REGISTRA SALDOS DE ESTOQUE
End Sub
Public Sub Processa_Escrituracao(cDtIni As String, cDtFim As String)
'DoCmd.setwarnings (False)
Call ConnectToDataBase


Call Calcula_Custo_Medio(cDtIni, cDtFim)

Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

'strSQL = ("UPDATE tbVendasDet SET tbVendasDet.VlrLiquido = tbVendasDet.ValorTot-tbVendasDet.Valor_ICMS-tbVendasDet.Valor_PIS-tbVendasDet.Valor_Cofins-tbVendasDet.Valor_IPI-tbVendasDet.Valor_ICMS_ST " & _
'"WHERE tbVendasDet.VlrLiquido=0;")
'Conn.Execute strSQL
'
''CALCULA CUSTO DE PRODUCAO
'
''CALCULA CUSTO DE PRODUCAO
'strSQL = ("update tbcadprod set CMed_Unit = 0, CMed_Tot = 0 where CMed_Unit = null;")
'Conn.Execute strSQL
'
'
''BUSCA O CUSTO UNITARIO DE PRODUTOS DE REVENDA COM TROCA DE CÓDIGO E GRANEL
'Set rst = Db.OpenRecordset("SELECT tbVendas.ANO, tbVendas.MES, tbVendas.DataEmissao, tbVendasDet.IDProd, tbVendasDet.Qnt, tbVendasDet.CustoMedio, tbVendasDet.VlrLiquido, tbCadProd.CMed_Unit, tbCadProd.DescProd, tbCadProd.LITROS, tbCadProd.IdProd_Revenda, tbcadprod_1.DescProd,tbcadprod_1.Litros as LitrosOriginal, tbcadprod_1.CMed_Unit as CMed_Unit_Original, tbcadprod_1.Estoque " & _
'"FROM (tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) LEFT JOIN tbcadprod AS tbcadprod_1 ON tbCadProd.IdProd_Revenda = tbcadprod_1.IDProd " & _
'"WHERE (((tbVendasDet.CustoMedio)=0) AND ((tbVendasDet.CFOP)<>'1604'));")
'
'Do Until rst.EOF
'On Error Resume Next
'
'strSQL = ("UPDATE tbCadProd SET tbCadProd.Estoque = 0, tbCadProd.CMed_Unit = " & Replace((rst!CMed_Unit_Original / rst!LitrosOriginal) * rst!Litros, ",", ".") & " where tbCadProd.IDProd = " & rst!IdProd & ";")
'
'
'Conn.Execute strSQL
'rst.MoveNext
'Loop
''BUSCA O CUSTO UNITARIO DE PRODUTOS DE REVENDA COM TROCA DE CÓDIGO E GRANEL
'
'
''REGISTRA CUSTO MEDIO DA VENDA E VALOR LIQUIDO
'strSQL = ("UPDATE tbVendasDet SET tbVendasDet.VlrLiquido = tbVendasDet.ValorTot-tbVendasDet.Valor_ICMS-tbVendasDet.Valor_PIS-tbVendasDet.Valor_Cofins-tbVendasDet.Valor_IPI-tbVendasDet.Valor_ICMS_ST " & _
'"WHERE (((tbVendasDet.VlrLiquido)=0));")
'Conn.Execute strSQL
'strSQL = ("UPDATE tbVendasDet SET tbVendasDet.VlrLiquido = tbVendasDet.ValorTot-tbVendasDet.Valor_ICMS-tbVendasDet.Valor_PIS-tbVendasDet.Valor_Cofins-tbVendasDet.Valor_IPI-tbVendasDet.Valor_ICMS_ST " & _
'"WHERE (((tbVendasDet.VlrLiquido) is null));")
'Conn.Execute strSQL
'
'strSQL = ("UPDATE tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd SET tbVendasDet.CustoMedio = tbCadProd.CMed_Unit*tbVendasDet.Qnt, tbVendasDet.Margem = (tbVendasDet.VlrLiquido-(tbCadProd.CMed_Unit*tbVendasDet.Qnt))/tbVendasDet.VlrLiquido " & _
'"WHERE (((tbVendasDet.CustoMedio) is null));")
'Conn.Execute strSQL
'strSQL = ("UPDATE tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd SET tbVendasDet.CustoMedio = tbCadProd.CMed_Unit*tbVendasDet.Qnt, tbVendasDet.Margem = (tbVendasDet.VlrLiquido-(tbCadProd.CMed_Unit*tbVendasDet.Qnt))/tbVendasDet.VlrLiquido " & _
'"WHERE (((tbVendasDet.CustoMedio) =0));")
'Conn.Execute strSQL
'strSQL = ("UPDATE tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd SET tbVendasDet.CustoMedio = tbCadProd.CMed_Unit*tbVendasDet.Qnt, tbVendasDet.Margem = (tbVendasDet.VlrLiquido-(tbCadProd.CMed_Unit*tbVendasDet.Qnt))/tbVendasDet.VlrLiquido " & _
'"WHERE (((tbVendasDet.CustoMedio)=null));")
'Conn.Execute strSQL
''REGISTRA CUSTO MEDIO DA VENDA E VALOR LIQUIDO
'
'
'
'
''CORRIGE CUSTO DE PRODUÇÃO DE PRODUTOS PRODUZIDOS
'Call Calc_CustoProducao
''CORRIGE CUSTO DE PRODUCAO DE PRODUTOS PRODUZIDOS
'Call ConnectToDataBase
''ajustes
''REGISTRA SALDOS DE ESTOQUE
'strSQL = ("UPDATE tbCadProd set Estoque = 0, CMed_Unit = 0, CMed_Tot=0 ")
'Conn.Execute strSQL
''entradas +
'Set rstSaldo = Db.OpenRecordset("SELECT tbComprasDet.IDProd, Sum(tbComprasDet.Qnt) AS QT, Sum(tbComprasDet.ValorTot) AS ValorTot FROM tbComprasDet GROUP BY tbComprasDet.IDProd;")
'    Do Until rstSaldo.EOF
'    strSQL = ("UPDATE tbCadProd set Estoque = " & Replace(rstSaldo!Qt, ",", ".") & ", CMed_Unit = " & Replace(rstSaldo!ValorTot / rstSaldo!Qt, ",", ".") & ", CMed_Tot = " & Replace(rstSaldo!ValorTot, ",", ".") & " where IdProd = " & rstSaldo!IdProd & "")
'    Conn.Execute strSQL
'    rstSaldo.MoveNext
'    Loop
''saidas Vendas -
'
'Set rstSaldo = Db.OpenRecordset("SELECT tbVendasDet.IDProd, Sum(tbVendasDet.Qnt) AS QT, Sum(tbVendasDet.CustoMedio) AS ValorTot FROM tbVendasDet where CFOP <> '1551' and CFOP <> '5920' GROUP BY tbVendasDet.IDProd ; ")
'    Do Until rstSaldo.EOF
''    strSQL = ("UPDATE tbCadProd set Estoque = Estoque - " & Replace(rstSaldo!Qt, ",", ".") & ", CMed_Tot = CMed_Tot - " & Replace(rstSaldo!ValorTot, ",", ".") & ", CMed_Unit = (CMed_Tot - " & Replace(rstSaldo!ValorTot, ",", ".") & " ) / (Estoque - " & Replace(rstSaldo!Qt, ",", ".") & ") where IdProd = " & rstSaldo!IDProd & "")
' '   Conn.Execute strSQL
'    rstSaldo.MoveNext
'    Loop
'    'DoCmd.setwarnings (False)
''ajustes manuais + ou -
''consumos e perdas diversas -
''producao +
'Set rstSaldo = Db.OpenRecordset("SELECT tb_Ajuste_Estoque.IDProd, tb_Ajuste_Estoque.QT, tb_Ajuste_Estoque.CMed_Tot, tb_Ajuste_Estoque.CMed_Unit FROM tb_Ajuste_Estoque order by Data Asc;")
'    Do Until rstSaldo.EOF
'    strSQL = ("UPDATE tbCadProd set Estoque = Estoque + " & Replace(rstSaldo!Qt, ",", ".") & ", CMed_Tot = CMed_Tot + " & Replace(rstSaldo!CMed_Tot, ",", ".") & " where IdProd = " & rstSaldo!IdProd & "")
'    Conn.Execute strSQL
'    ', CMed_Unit = (CMed_Tot + " & Replace(rstSaldo!ValorTot, ",", ".") & " ) / (Estoque + " & Replace(rstSaldo!QT, ",", ".") & ")
'    strSQL = ("UPDATE tbCadProd set CMed_Unit = CMed_Tot / Estoque where IdProd = " & rstSaldo!IdProd & " and Estoque > 0 ")
'    Conn.Execute strSQL
'    strSQL = ("UPDATE tbCadProd set CMed_Unit = " & Replace(rstSaldo!CMed_Unit * -1, ",", ".") & " where IdProd = " & rstSaldo!IdProd & " and Estoque < 0 ")
'    Conn.Execute strSQL
'    strSQL = ("UPDATE tbCadProd set CMed_Unit = " & Replace(rstSaldo!CMed_Unit * -1, ",", ".") & ", CMed_Tot = 0 where IdProd = " & rstSaldo!IdProd & " and Estoque = 0 ")
'    Conn.Execute strSQL
'
'    rstSaldo.MoveNext
'    Loop
'
''corrigir negativos unitarios estoque zero
'strSQL = ("update tbCadProd set CMed_Unit = CMed_Unit * -1 where Estoque = 0 and CMed_Unit < 0")
'Conn.Execute strSQL
'
''zerar saldos de energia
'strSQL = ("UPDATE tbCadProd set Estoque = 0, CMed_Tot=0 where idProd = 2800")
'Conn.Execute strSQL
''REGISTRA SALDOS DE ESTOQUE


'CORRIGIR O PIS COFINS DA ENERGIA ELÉTRICA PARA A ALIQUOTA BASICA
'1,65% PIS E 7,6% COFINS - Somente para LUCRO REAL
strSQL = ("UPDATE tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra SET tbComprasDet.Aliq_PIS = 1.65, tbComprasDet.Aliq_Cofins = 7.6, tbComprasDet.Valor_PIS = tbComprasDet.BaseCalculo*0.0165, tbComprasDet.Valor_Cofins = tbComprasDet.BaseCalculo*0.076 " & _
"WHERE (((tbCompras.IdFornecedor)=1131) AND ((tbCompras.ANO) >=2020)); ")
Conn.Execute strSQL
'CORRIGIR PIS COFINS LUCRO REAL TB NA TABELA ENERGIA
strSQL = ("update tbEnergia set Aliq_PIS = 1.65, AliqCofins = 7.6, VlPIS = round(VlBaseCalcICMS * (Aliq_PIS/100),2), VlCofins =  round(VlBaseCalcICMS * (AliqCofins/100),2) where ANO >= 2020 AND VlPIS is null and VlCofins is null ;")
Conn.Execute strSQL
'COMECEI LUCRO REAL EM JAN 2020


'REGISTRA ENERGIA ELÉTRICA COMO UMA COMPRA
'HEADER
strSQL = ("INSERT INTO tbCompras ( IdFornecedor, TipoNF, NatOperacao, DataEmissao, NumNF, Serie, ChaveNF, VlrTotalProdutos, VlrTotalFrete, VlrTotalSeguro, VlrDesconto, VlrDespesas, ICMS_BaseCalc, ICMS_Valor, ICMS_ST_BaseCalc, ICMS_ST_Valor, IPI_Valor, PIS_Valor, COFINS_Valor, VlrTOTALNF ) " & _
"SELECT tbEnergia.IdFor, '1-SAIDA' AS TIPO_NF, 'VENDA DE ENERGIA ELETRICA NO ESTADO' AS NatOp, tbEnergia.DataNota, tbEnergia.NumNota, tbEnergia.Serie, tbEnergia.ChaveNota, tbEnergia.ValorTotal, 0 AS Expr1, 0 AS Expr2, tbEnergia.ValorTotalDesc, 0 AS Expr3, tbEnergia.VlBaseCalcICMS, tbEnergia.VlICMS, 0 AS Expr4, 0 AS Expr5, 0 AS Expr6, tbEnergia.VlPIS, tbEnergia.VlCofins, tbEnergia.ValorTotal " & _
"FROM tbEnergia LEFT JOIN tbCompras ON tbEnergia.ChaveNota = tbCompras.ChaveNF " & _
"WHERE (((tbCompras.ChaveNF) Is Null));")
Conn.Execute strSQL

'LINE
strSQL = ("INSERT INTO tbComprasDet ( IDCompra, IDProd, Qnt, ValorUnit, ValorTot, VlrFrete, VlrSeguro, VlrDesc, VlrOutro, CFOP, CFOP_DESC, Pedido, Cd_Origem, Origem, CST, CST_DESC, BaseCalculo, Aliq_ICMS, Valor_ICMS, Aliq_PIS, Valor_PIS, Aliq_Cofins, Valor_Cofins, Aliq_IPI, Valor_IPI, MVA_ST, Aliq_ICMS_ST, BaseCalc_ST, Valor_ICMS_ST, InfoAdicional, CFOP_ESCRITURADA, CFOP_ESC_DESC, LancFiscal ) " & _
"SELECT tbCompras.ID, 2800 AS Expr1, tbEnergia.TotalKWH, VlrTotalProdutos/TotalKWH AS Expr2, tbCompras.VlrTotalProdutos, 0 AS Expr3, 0 AS Expr4, tbCompras.VlrDesconto, 0 AS Expr5, 5253 AS Expr6, 'Venda de energia elétrica para estabelecimento comercial' AS Expr7, 0 AS Expr8, '0', 'NACIONAL' AS Expr9, '00' AS Expr10, 'Tributada integralmente' AS Expr11, tbEnergia.VlBaseCalcICMS, tbEnergia.Aliq_ICMS, tbEnergia.VlICMS, tbEnergia.Aliq_PIS, tbEnergia.VlPIS, tbEnergia.AliqCofins, tbEnergia.VlCofins, 0 AS Expr12, 0 AS Expr13, 0 AS Expr14, 0 AS Expr15, 0 AS Expr16, 0 AS Expr17, '' AS Expr18, 1252 AS Expr19, 'Compra de energia elétrica por estabelecimento industrial' AS Expr20, 'CREDITO' AS Expr21 " & _
"FROM (tbCompras LEFT JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) INNER JOIN tbEnergia ON tbCompras.ChaveNF = tbEnergia.ChaveNota " & _
"WHERE (((tbComprasDet.IDCompra) Is Null));")
Conn.Execute strSQL

'REGISTRA ALUGUEL COMO UMA COMPRA
strSQL = ("UPDATE tb_Aluguel SET tb_Aluguel.ID_Chave = CONCAT('ALUGUEL-', ID);")
Conn.Execute strSQL

strSQL = ("INSERT INTO tbCompras ( IdFornecedor, TipoNF, NatOperacao, DataEmissao, VlrTotalProdutos, VlrTotalFrete, VlrTotalSeguro, VlrDesconto, VlrDespesas, ICMS_BaseCalc, ICMS_Valor, ICMS_ST_BaseCalc, ICMS_ST_Valor, IPI_Valor, PIS_Valor, COFINS_Valor, VlrTOTALNF, ChaveNF, ANO, MES ) " & _
"SELECT tb_Aluguel.Id_For, '1-SAIDA' AS 'TIPO NF', 'LOCACAO IMOVEL' AS NatOp, tb_Aluguel.Data_Fim, tb_Aluguel.Valor_Aluguel, 0 AS Expr1, 0 AS Expr2, 0 AS Expr7, 0 AS Expr3, 0 AS Expr8, 0 AS Expr9, 0 AS Expr4, 0 AS Expr5, 0 AS Expr6, 0 AS Expr10, 0 AS Expr11, tb_Aluguel.Valor_Aluguel, tb_Aluguel.ID_Chave, Year(Data_Fim) AS ANO, Month(Data_Fim) AS MES " & _
"FROM tb_Aluguel LEFT JOIN tbCompras ON tb_Aluguel.id_chave = tbCompras.ChaveNF " & _
"WHERE (((tb_Aluguel.Valor_Aluguel)>0) AND ((tbCompras.ChaveNF) Is Null) AND ((tb_Aluguel.Data_Pagamento) Is Not Null));")
Conn.Execute strSQL
'REGISTRA ALUGUEL COMO UMA COMPRA


'REGISTRA O CONTAS A PAGAR
'ENERGIA
strSQL = ("INSERT INTO tb_Detalhe_Boletos_Compras ( DtEmissao, DtVencimento, DtPagamento, Id_Fornecedor, ValorOriginal, ValorPago, Chave_Nfe, STATUS, Fornecedor ) " & _
             "SELECT tbCompras.DataEmissao, tbCompras.DataEmissao, tbCompras.DataEmissao, tbCompras.IdFornecedor, tbCompras.VlrTOTALNF, tbCompras.VlrTOTALNF, tbCompras.ChaveNF, 'PAGO' AS Expr1, tbFornecedor.RazaoSocial " & _
             "FROM tbFornecedor INNER JOIN (tb_Detalhe_Boletos_Compras RIGHT JOIN tbCompras ON tb_Detalhe_Boletos_Compras.Chave_Nfe = tbCompras.ChaveNF) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
             "WHERE (((tbCompras.IdFornecedor)=1131) AND ((tb_Detalhe_Boletos_Compras.Chave_Nfe) Is Null));")
Conn.Execute strSQL

'ALUGUEL
strSQL = ("INSERT INTO tb_Detalhe_Boletos_Compras ( DtEmissao, DtVencimento, DtPagamento, Id_Fornecedor, ValorOriginal, ValorPago, Chave_Nfe, STATUS, Fornecedor ) " & _
            "SELECT tbCompras.DataEmissao, tbCompras.DataEmissao, tbCompras.DataEmissao, tbCompras.IdFornecedor, tbCompras.VlrTOTALNF, tbCompras.VlrTOTALNF, tbCompras.ChaveNF, 'PAGO' AS Expr1, tbFornecedor.RazaoSocial " & _
            "FROM tbFornecedor INNER JOIN (tb_Detalhe_Boletos_Compras RIGHT JOIN tbCompras ON tb_Detalhe_Boletos_Compras.Chave_Nfe = tbCompras.ChaveNF) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
            "WHERE (((tbCompras.NatOperacao)='LOCACAO IMOVEL') AND ((tb_Detalhe_Boletos_Compras.Chave_Nfe) Is Null));")
Conn.Execute strSQL

'REGISTRA O CONTAS A PAGAR



'atualiza ano e mes da energia
strSQL = ("UPDATE tbCompras SET tbCompras.ANO = Year(tbCompras.DataEmissao), tbCompras.MES = Month(tbCompras.DataEmissao) " & _
"WHERE (((tbCompras.ANO) Is Null) AND ((tbCompras.MES) Is Null) AND ((tbCompras.IdFornecedor)=1131));")
Conn.Execute strSQL



'CALCULAR ICMS DE FORNECEDORES SIMPLES QUE INFORMARAM NO CAMPO OBSERVAÇÃO DA NF O VALOR DO ICMS PARA CREDITO
'AQUI FAZ LINE E HEADER
strSQL = ("UPDATE (tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor) INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra SET tbComprasDet.Aliq_ICMS = tbFornecedor.Aliq_SIMPLES, tbComprasDet.Valor_ICMS = tbComprasDet.ValorTot*(tbFornecedor.Aliq_SIMPLES/100), tbCompras.ICMS_Valor = tbCompras.VlrTotalProdutos*(tbFornecedor.Aliq_SIMPLES/100) " & _
"WHERE (((tbFornecedor.Tipo)='FORNECEDOR') AND ((tbFornecedor.CRT)='SIMPLES NACIONAL') AND ((tbFornecedor.Aliq_SIMPLES) Is Not Null) AND ((tbComprasDet.Aliq_ICMS)=0) AND ((tbComprasDet.Valor_ICMS)=0) AND ((tbCompras.ICMS_Valor)=0));")
Conn.Execute strSQL


'CORRIGIR QUANDO O FORNECEDOR NÃO INFORMA A BASE DE CALCULO DO ICMS NO XML
'CRIA TEMP
strSQL = ("delete from TEMP_CORRIGE_BASE_CALC;")
Conn.Execute strSQL

strSQL = ("" & _
"insert into TEMP_CORRIGE_BASE_CALC " & _
"SELECT tbComprasDet.IDCompra, Sum(tbComprasDet.BaseCalculo) AS SomaDeBaseCalculo, tbCompras.ICMS_BaseCalc " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"GROUP BY tbComprasDet.IDCompra, tbCompras.ICMS_BaseCalc " & _
"HAVING (((Sum(tbComprasDet.BaseCalculo))>0) AND ((tbCompras.ICMS_BaseCalc)=0));")
Conn.Execute strSQL

'ATUALIZA TB COMPRAS HEADER
strSQL = ("" & _
"UPDATE TEMP_CORRIGE_BASE_CALC INNER JOIN tbCompras ON TEMP_CORRIGE_BASE_CALC.IDCompra = tbCompras.ID SET tbCompras.ICMS_BaseCalc = TEMP_CORRIGE_BASE_CALC.SomaDeBaseCalculo " & _
"WHERE (((tbCompras.ICMS_BaseCalc)=0));")
Conn.Execute strSQL




'ENTRADAS

'1000    ENTRADAS OU AQUISIÇÕES DE SERVIÇOS DO ESTADO
'1101    Compra para industrialização
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1101', tbComprasDet.LancFiscal = 'CREDITO' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.MAT_PRIMA)='SIM'));")
Conn.Execute strSQL
'1101    Compra para industrialização - Embalagem
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1101', tbComprasDet.LancFiscal = 'CREDITO' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.EMBALAGEM)='SIM'));")
Conn.Execute strSQL
'1102    Compra para comercialização - Revenda
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1102', tbComprasDet.LancFiscal = 'REVENDA' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.REVENDA)='SIM'));")
Conn.Execute strSQL
'1405    Compra para comercialização - ICMS ST
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1405', tbComprasDet.LancFiscal = 'SUBS_TRIB' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.REVENDA)='SIM')) and CFOP = '5405';")
Conn.Execute strSQL
'1252    Compra de energia elétrica por estabelecimento industrial
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1252', tbComprasDet.LancFiscal = 'CREDITO' WHERE (((tbFornecedor.UF)='SP') AND ((tbFornecedor.IDFor)=1131));")
Conn.Execute strSQL


'CFOP ESCRITURADA JÁ LANÇADA NO REGISTRO DA COMPRA MANUAL

'1352    Aquisição de serviço de transporte por estabelecimento industrial
'Ja faz esse lançamento em algum outro lugar - Tabela OK

'1556    Compra de material para uso ou consumo
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1556', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.CONSUMO)='SIM'));")
Conn.Execute strSQL

strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1556', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.MAT_PUBLICIDADE)='SIM'));")
Conn.Execute strSQL
'1551    Compra de bem para o ativo imobilizado
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1551', tbComprasDet.LancFiscal = 'IMOBILIZADO' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.IMOBILIZADO)='SIM'));")
Conn.Execute strSQL
'1911    Entrada de amostra grátis
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1911', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)='SP') AND ((tbComprasDet.CFOP)='5911'));")
Conn.Execute strSQL
'1911    Compra de material de escritório
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1911', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.MAT_ESCRITORIO)='SIM'));")
Conn.Execute strSQL
'1911    Compra de instalações comerciais
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1911', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)='SP') AND ((tbCadProd.INST_COMERCIAIS)='SIM'));")
Conn.Execute strSQL
'1949    Outra entrada de mercadoria ou presta??o de servi?o n?o especificada
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1949', tbComprasDet.LancFiscal = 'OUTROS' WHERE tbFornecedor.UF='SP' AND CFOP_ESCRITURADA is NULL;")
Conn.Execute strSQL






'2000    ENTRADAS OU AQUISIÇÕES DE SERVIÇOS DE OUTROS ESTADOS
'2101    Compra para industrialização
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2101', tbComprasDet.LancFiscal = 'CREDITO' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.MAT_PRIMA)='SIM'));")
Conn.Execute strSQL
'2101    Compra para industrialização - Embalagem
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2101', tbComprasDet.LancFiscal = 'CREDITO' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.EMBALAGEM)='SIM'));")
Conn.Execute strSQL
'2102    Compra para comercialização
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2102', tbComprasDet.LancFiscal = 'REVENDA' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.REVENDA)='SIM'));")
Conn.Execute strSQL
'2102    Compra para comercialização - ICMS ST
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '1102', tbComprasDet.LancFiscal = 'SUBS_TRIB' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.REVENDA)='SIM')) AND tbComprasDet.CST = 10;")
Conn.Execute strSQL

'2252    Compra de energia elétrica por estabelecimento industrial
'2352    Aquisição de serviço de transporte por estabelecimento industrial

'2556    Compra de material para uso ou consumo
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2556', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.CONSUMO)='SIM'));")
Conn.Execute strSQL
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2556', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.MAT_PUBLICIDADE)='SIM'));")
Conn.Execute strSQL
'2551    Compra de bem para o ativo imobilizado
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2551', tbComprasDet.LancFiscal = 'IMOBILIZADO'  WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.IMOBILIZADO)='SIM'));")
Conn.Execute strSQL
'2911    Entrada de amostra grátis
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2911', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbComprasDet.CFOP)='6911'));")
Conn.Execute strSQL
'2911    Compra de material de escritório
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2911', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.MAT_ESCRITORIO)='SIM'));")
Conn.Execute strSQL
'2911    Compra de material de escritório
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2911', tbComprasDet.LancFiscal = 'OUTROS' WHERE (((tbFornecedor.UF)<>'SP') AND ((tbCadProd.INST_COMERCIAIS)='SIM'));")
Conn.Execute strSQL

'2949    Outra entrada de mercadoria ou presta??o de servi?o n?o especificada
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CFOP_ESCRITURADA = '2949', tbComprasDet.LancFiscal = 'OUTROS' WHERE tbFornecedor.UF<>'SP' AND CFOP_ESCRITURADA is NULL;")
Conn.Execute strSQL




'zera o valor referente a ICMS ST quando não há direito a crédito vira despesa
strSQL = ("UPDATE tbComprasDet INNER JOIN tbCadProd on tbComprasDet.IDProd = tbCadProd.IDProd set tbComprasDet.Valor_ICMS_ST = 0 where LancFiscal = 'SUBS_TRIB'  and REVENDA = 'SIM';")
Conn.Execute strSQL
strSQL = ("UPDATE tbComprasDet INNER JOIN tbCadProd on tbComprasDet.IDProd = tbCadProd.IDProd set tbComprasDet.Valor_ICMS_ST = 0 where LancFiscal = 'REVENDA'  and REVENDA = 'SIM';")
Conn.Execute strSQL

'Credito do IPI de comerciante que não recolhe IPI de material usado no proceso produtivo credita 50%
'Ao adquirir mercadoria de comerciante (não-contribuinte de IPI), para ser aplicada no processo produtivo podemos creditar 50% do valor do IPI que teríamos direito se adquirido de contribuinte desse imposto. (Fonte: artigo 165 do RIPI – Decreto 4.544/02)
strSQL = ("UPDATE (tbFornecedor INNER JOIN (tbCadProd INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbFornecedor.IDFor = tbCompras.IdFornecedor) INNER JOIN tbIPI ON tbCadProd.NCM = tbIPI.NCM SET tbComprasDet.Valor_IPI = round(BaseCalculo*(tbIPI.ALIQ_IPI/100)/2,4), tbComprasDet.Aliq_IPI = tbIPI.ALIQ_IPI " & _
"WHERE (((tbCompras.DataEmissao)>='2017-7-1') AND ((tbComprasDet.Valor_IPI)=0) AND ((tbCadProd.MAT_PRIMA)='SIM')) OR (((tbCompras.DataEmissao)>=7/1/2017) AND ((tbComprasDet.Valor_IPI)=0) AND ((tbCadProd.EMBALAGEM)='SIM'));")
Conn.Execute strSQL

'tem que atualizar o header do pedido
strSQL = ("delete from tbCorrecaoIPI_temp;")
Conn.Execute strSQL

strSQL = ("INSERT INTO tbCorrecaoIPI_temp ( ID, IPI_TOT ) " & _
"SELECT tbCompras.ID, Sum(tbComprasDet.Valor_IPI) AS IPI_TOT " & _
"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"GROUP BY tbCompras.ID, tbCompras.DataEmissao, tbCadProd.MAT_PRIMA, tbCadProd.EMBALAGEM " & _
"HAVING (((tbCompras.DataEmissao)>='2017-7-1') AND ((tbCadProd.MAT_PRIMA)='SIM') AND ((Sum(tbComprasDet.Valor_IPI))>0)) OR (((tbCompras.DataEmissao)>='2017-7-1') AND ((tbCadProd.EMBALAGEM)='SIM') AND ((Sum(tbComprasDet.Valor_IPI))>0));")
Conn.Execute strSQL

strSQL = ("UPDATE tbCorrecaoIPI_temp INNER JOIN tbCompras ON tbCorrecaoIPI_temp.ID = tbCompras.ID SET tbCompras.IPI_Valor = tbCorrecaoIPI_temp.IPI_TOT;")
Conn.Execute strSQL

''Credito de ICMS de Imobilizado com ICMS ST
'O crédito do ICMS será calculado sobre o valor da base de cálculo que seria devida pelo fornecedor, caso a mercadoria estivesse submetida ao regime comum de tributação, conforme artigo 272 do RICMS/SP.
'Artigo 272 - O contribuinte que receber, com imposto retido, mercadoria não destinada a comercialização subsequente, aproveitará o crédito fiscal, quando admitido, calculando-o mediante aplicação da alíquota interna sobre a base de cálculo que seria atribuída à operação própria do remetente, caso estivesse submetida ao regime comum de tributação (Lei 6.374/89, art. 36, com alteração da Lei 9.359/96, art. 2º, I).
strSQL = ("UPDATE tbCadProd INNER JOIN ((tbICMS_ALIQ_Basica_UF INNER JOIN (tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor) ON tbICMS_ALIQ_Basica_UF.UF = tbFornecedor.UF) INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd) SET tbComprasDet.Aliq_ICMS = tbICMS_ALIQ_Basica_UF.ALIQ_ICMS, tbComprasDet.Valor_ICMS = tbComprasDet.BaseCalculo*(tbICMS_ALIQ_Basica_UF.ALIQ_ICMS/100) " & _
"WHERE (((tbComprasDet.CST)='60') AND ((tbCadProd.IMOBILIZADO)='SIM') AND ((tbComprasDet.Valor_ICMS)=0)); ")
Conn.Execute strSQL

'Tem que Atualizar o Header da compra
Set rstCompras_Imob_ST = Db.OpenRecordset("SELECT tbComprasDet.IDCompra, tbComprasDet.CST, Sum(tbComprasDet.Valor_ICMS) AS Valor_ICMS, tbCadProd.IMOBILIZADO " & _
"FROM tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd) " & _
"GROUP BY tbComprasDet.IDCompra, tbComprasDet.CST, tbCadProd.IMOBILIZADO " & _
"HAVING (((tbComprasDet.CST)='60') AND ((tbCadProd.IMOBILIZADO)='SIM'));")
Do Until rstCompras_Imob_ST.EOF
    strSQL = ("update tbCompras set ICMS_Valor = " & Replace(rstCompras_Imob_ST!Valor_ICMS, ",", ".") & " where ID = " & rstCompras_Imob_ST!IDCompra & " and ICMS_Valor < " & Replace(rstCompras_Imob_ST!Valor_ICMS, ",", ".") & "")
    Conn.Execute strSQL
    rstCompras_Imob_ST.MoveNext
Loop



'SAÍDAS
'5000    SAÍDAS OU PRESTAÇÕES DE SERVIÇOS PARA O ESTADO
'5101    Venda de produção do estabelecimento
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '5101', tbVendasDet.LancFiscal = 'DEBITO' WHERE (((tbCliente.UF)='SP') AND ((tbCadProd.PROD_FINAL)='SIM'));")
Conn.Execute strSQL
'5102    Venda de mercadoria adquirida ou recebida de terceiros
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '5102',tbVendasDet.LancFiscal = 'DEBITO' WHERE (((tbCliente.UF)='SP') AND ((tbCadProd.REVENDA)='SIM'));")
Conn.Execute strSQL
'5405    Venda de mercadoria adquirida ou recebida de terceiros em operação com mercadoria sujeita ao regime de substituição tributária, na condição de contribuinte substituído
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '5405',tbVendasDet.LancFiscal = 'DEBITO' WHERE (((tbCliente.UF)='SP') AND ((tbCadProd.REVENDA)='SIM')) and Aliq_ICMS = 0;")
Conn.Execute strSQL
'5920    Remessa de vasilhame ou sacaria - EMBALAGEM SIMPLES REMESSA
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '5920',tbVendasDet.LancFiscal = 'OUTROS' WHERE (((tbCliente.UF)='SP') AND ((tbCadProd.EMBALAGEM)='SIM') AND ((tbVendasDet.CFOP)='5920'));")
Conn.Execute strSQL


'6000    SAÍDAS OU PRESTAÇÕES DE SERVIÇOS PARA OUTROS ESTADOS
'6101    Venda de produção do estabelecimento
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '6101', tbVendasDet.LancFiscal = 'DEBITO' WHERE (((tbCliente.UF)<>'SP') AND ((tbCadProd.PROD_FINAL)='SIM'));")
Conn.Execute strSQL
'6102    Venda de mercadoria adquirida ou recebida de terceiros
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '6102', tbVendasDet.LancFiscal = 'DEBITO' WHERE (((tbCliente.UF)<>'SP') AND ((tbCadProd.REVENDA)='SIM'));")
Conn.Execute strSQL
'6405    Venda de mercadoria adquirida ou recebida de terceiros em operação com mercadoria sujeita ao regime de substituição tributária, na condição de contribuinte substituído
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '6405',tbVendasDet.LancFiscal = 'DEBITO' WHERE (((tbCliente.UF)<>'SP') AND ((tbCadProd.REVENDA)='SIM')) and Aliq_ICMS = 0;")
Conn.Execute strSQL
'6920    Remessa de vasilhame ou sacaria - EMBALAGEM SIMPLES REMESSA
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '5920', tbVendasDet.LancFiscal = 'OUTROS' WHERE (((tbCliente.UF)<>'SP') AND ((tbCadProd.EMBALAGEM)='SIM') AND ((tbVendasDet.CFOP)='5920'));")
Conn.Execute strSQL

'ENTRADA CIAP
'1604 - Lançamento do crédito relativo à compra de bem para o ativo imobilizado (mês anterior, nota de entrada)
strSQL = ("UPDATE (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CFOP_ESCRITURADA = '1604',tbVendasDet.LancFiscal = 'OUTROS' WHERE (((tbVendas.TipoNF)='0-ENTRADA') AND ((tbVendas.IdCliente)=832) AND ((tbVendasDet.CFOP)='1604'));")
Conn.Execute strSQL



'TRANSPORTES
'ENTRADAS
'1350    AQUISIÇÕES DE SERVIÇOS DE TRANSPORTE NO ESTADO

'1352    Aquisição de serviço de transporte por estabelecimento industrial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '1352' WHERE (((tbTransportes.DestinatarioCNPJ)='23866944000141') AND ((tbTransportes.RemetenteUF)='SP') AND ((tbTransportes.DataEmissao)>='2017-7-1'));")
Conn.Execute strSQL

'1353    Aquisição de serviço de transporte por estabelecimento comercial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '1353' WHERE (((tbTransportes.DestinatarioCNPJ)='23866944000141') AND ((tbTransportes.RemetenteUF)='SP') AND ((tbTransportes.DataEmissao)<'2017-7-1'));")
Conn.Execute strSQL
'2350    AQUISIÇÕES DE SERVIÇOS DE TRANSPORTE FORA DO ESTADO

'2352    Aquisição de serviço de transporte por estabelecimento industrial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '2352' WHERE (((tbTransportes.DestinatarioCNPJ)='23866944000141') AND ((tbTransportes.RemetenteUF)<>'SP') AND ((tbTransportes.DataEmissao)>='2017-7-1'));")
Conn.Execute strSQL

'2353    Aquisição de serviço de transporte por estabelecimento comercial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '2353' WHERE (((tbTransportes.DestinatarioCNPJ)='23866944000141') AND ((tbTransportes.RemetenteUF)<>'SP') AND ((tbTransportes.DataEmissao)<'2017-7-1'));")
Conn.Execute strSQL


'TRANSPORTES
'SAIDAS

'5350    PRESTAÇÕES DE SERVIÇOS DE TRANSPORTE NO ESTADO
'5352    Prestação de serviço de transporte a estabelecimento industrial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '5352' WHERE (((tbTransportes.RemetenteCNPJ)='23866944000141') AND ((tbTransportes.DestinatarioUF)='SP') AND ((tbTransportes.DataEmissao)>='2017-7-1'));")
Conn.Execute strSQL

'5353    Prestação de serviço de transporte a estabelecimento comercial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '5353' WHERE (((tbTransportes.RemetenteCNPJ)='23866944000141') AND ((tbTransportes.DestinatarioUF)='SP') AND ((tbTransportes.DataEmissao)<'2017-7-1'));")
Conn.Execute strSQL

'6350    PRESTAÇÕES DE SERVIÇOS DE TRANSPORTE FORA DO ESTADO
'6352    Prestação de serviço de transporte a estabelecimento industrial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '6352' WHERE (((tbTransportes.RemetenteCNPJ)='23866944000141') AND ((tbTransportes.DestinatarioUF)<>'SP') AND ((tbTransportes.DataEmissao)>='2017-7-1'));")
Conn.Execute strSQL

'6353    Prestação de serviço de transporte a estabelecimento comercial
strSQL = ("UPDATE tbTransportes SET tbTransportes.CFOP_ESCRITURADA = '6353' WHERE (((tbTransportes.RemetenteCNPJ)='23866944000141') AND ((tbTransportes.DestinatarioUF)<>'SP') AND ((tbTransportes.DataEmissao)<'2017-7-1'));")
Conn.Execute strSQL


'AVALIA SE O TRANSPORTE É CREDITAVEL
strSQL = ("UPDATE tbTransportes SET tbTransportes.Creditavel = 'NAO', tbTransportes.LancFiscal = 'OUTROS';")
Conn.Execute strSQL

'QUANDO EU SOU TOMADOR E DESTINATARIO
strSQL = ("UPDATE tbTransportes SET tbTransportes.Creditavel = 'SIM', tbTransportes.LancFiscal = 'CREDITO' WHERE (((tbTransportes.DestinatarioCNPJ)='23866944000141') AND ((tbTransportes.Tomador)='DESTINATARIO'));")
Conn.Execute strSQL

'QUANDO EU SOU TOMADOR E REMETENTE
strSQL = ("UPDATE tbTransportes SET tbTransportes.Creditavel = 'SIM', tbTransportes.LancFiscal = 'CREDITO' WHERE (((tbTransportes.RemetenteCNPJ)='23866944000141') AND ((tbTransportes.Tomador)='REMETENTE'));")
Conn.Execute strSQL


'REGISTRA CÓDIGOS DE CST DE COMPRAS
'CST ICMS
'FORNECEDOR SIMPLES SEM CREDITO DE IMPOSTOS
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_ICMS = '102' " & _
"WHERE (((tbFornecedor.CRT)='SIMPLES NACIONAL'));")
Conn.Execute strSQL

'CSTS ICMS REGIME NORMAL BASEADO NO CST INFORMADO PELO FORNECEDOR NO XML
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_ICMS = CONCAT(tbComprasDet.Cd_Origem,CST) " & _
"WHERE (((tbFornecedor.CRT)<>'SIMPLES NACIONAL'));")
Conn.Execute strSQL

'CST IPI
'CST 00 Entrada com recuperação de crédito
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_IPI = '00' WHERE (((tbComprasDet.LancFiscal)='CREDITO') AND ((tbFornecedor.CRT)='REGIME NORMAL'));")
Conn.Execute strSQL

'CST 49 Outras entradas - Simples Nacional
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_IPI = '49' WHERE (((tbFornecedor.CRT)='SIMPLES NACIONAL'));")
Conn.Execute strSQL

'Todas as outras entradas
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_IPI = '49' WHERE (((tbComprasDet.LancFiscal)='OUTROS'));")
Conn.Execute strSQL

'tratar o imobilizado
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_IPI = '00' WHERE (((tbComprasDet.LancFiscal)='IMOBILIZADO') AND ((tbComprasDet.CST_ICMS)<>'102'));")
Conn.Execute strSQL

strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_IPI = '49' WHERE (((tbComprasDet.LancFiscal)='IMOBILIZADO') AND ((tbComprasDet.CST_ICMS)='102'));")
Conn.Execute strSQL

'CST PIS

strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_PIS = '50' WHERE (((tbComprasDet.LancFiscal)='CREDITO'));")
Conn.Execute strSQL
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_PIS = '70' WHERE (((tbComprasDet.LancFiscal)='IMOBILIZADO'));")
Conn.Execute strSQL
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_PIS = '70' WHERE (((tbComprasDet.CST_PIS) Is Null) AND ((tbComprasDet.LancFiscal)='OUTROS'));")
Conn.Execute strSQL
'CST COFINS
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_Cofins = '50' WHERE (((tbComprasDet.LancFiscal)='CREDITO'));")
Conn.Execute strSQL
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_Cofins = '70' WHERE (((tbComprasDet.LancFiscal)='IMOBILIZADO'));")
Conn.Execute strSQL
strSQL = ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.CST_Cofins = '70' WHERE (((tbComprasDet.CST_Cofins) Is Null) AND ((tbComprasDet.LancFiscal)='OUTROS'));")
Conn.Execute strSQL
'tratar o CIAP
'CST IPI
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_IPI = '49' WHERE tbVendasDet.LancFiscal='OUTROS' AND tbVendasDet.CFOP_ESCRITURADA='1604';")
Conn.Execute strSQL
'CST PIS
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_PIS = '98' WHERE tbVendasDet.LancFiscal='OUTROS' AND tbVendasDet.CFOP_ESCRITURADA='1604';")
Conn.Execute strSQL
'CST COFINS
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_COFINS = '98' WHERE tbVendasDet.LancFiscal='OUTROS' AND tbVendasDet.CFOP_ESCRITURADA='1604';")
Conn.Execute strSQL


'BUSCA DESCRICAO CFOP
strSQL = ("UPDATE tbCFOP_3 INNER JOIN tbComprasDet ON tbCFOP_3.NIV3 = tbComprasDet.CFOP_ESCRITURADA SET tbComprasDet.CFOP_ESC_DESC = tbCFOP_3.DESC;")
Conn.Execute strSQL
strSQL = ("UPDATE tbCFOP_3 INNER JOIN tbVendasDet ON tbCFOP_3.NIV3 = tbVendasDet.CFOP_ESCRITURADA SET tbVendasDet.CFOP_ESC_DESC = tbCFOP_3.DESC;")
Conn.Execute strSQL
strSQL = ("UPDATE tbCFOP_3 INNER JOIN tbTransportes ON tbCFOP_3.NIV3 = tbTransportes.CFOP_ESCRITURADA SET tbTransportes.CFOP_ESC_DESC = tbCFOP_3.DESC;")
Conn.Execute strSQL




'REGISTRA CÓDIGOS CST DE VENDAS
'CST ICMS NORMAL
'000-Nacional, exceto as indicadas nos códigos 3, 4, 5 e 8 da Tabela A - Tributada integralmente
'CST ICMS ST
'010-Nacional, exceto as indicadas nos códigos 3, 4, 5 e 8 da Tabela A - Tributada e com cobrança do ICMS por substituição tributária
'POR ENQUANTO SEM ST ATÉ ATUALIZAR NA SEFAZ
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_ICMS = '000' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_ICMS) Is Null));")
Conn.Execute strSQL


'CST IPI
'50-saida tributada
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_IPI = '50' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_IPI) Is Null) AND ((tbVendasDet.CFOP)='5101' Or (tbVendasDet.CFOP)='6101') AND ((tbVendas.TipoNF)='1-SAIDA'));")
Conn.Execute strSQL

'53-saida NÃO tributada (Adquirida de terceiros industrial)
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_IPI = '53' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_IPI) Is Null) AND ((tbVendasDet.CFOP)='5102' Or (tbVendasDet.CFOP)='6102') AND ((tbVendas.TipoNF)='1-SAIDA'));")
Conn.Execute strSQL

'53-saida NÃO tributada (Adquirida de terceiros industrial)
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_IPI = '53' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_IPI) Is Null) AND ((tbVendasDet.CFOP)='5405' Or (tbVendasDet.CFOP)='6102') AND ((tbVendas.TipoNF)='1-SAIDA'));")
Conn.Execute strSQL



'CST PIS
'Produção do estabelecimento
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_PIS = '02' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_PIS) Is Null) AND ((tbVendasDet.CFOP)='5101' Or (tbVendasDet.CFOP)='6101') AND ((tbVendas.TipoNF)='1-SAIDA'));")
Conn.Execute strSQL

'Adiquirida de terceiros
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_PIS = '02' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_PIS) Is Null) AND ((tbVendasDet.CFOP)='5102' Or (tbVendasDet.CFOP)='6102') AND ((tbVendas.TipoNF)='1-SAIDA'));")
Conn.Execute strSQL


'CST COFINS
'Produção do Estabelecimento
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_COFINS = '02' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_COFINS) Is Null) AND ((tbVendasDet.CFOP)='5101' Or (tbVendasDet.CFOP)='6101') AND ((tbVendas.TipoNF)='1-SAIDA'));")
Conn.Execute strSQL

'Adiquirida de terceiros
strSQL = ("UPDATE tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda SET tbVendasDet.CST_COFINS = '02' " & _
"WHERE (((tbVendas.DataEmissao)>='2017-7-1') AND ((tbVendasDet.CST_COFINS) Is Null) AND ((tbVendasDet.CFOP)='5102' Or (tbVendasDet.CFOP)='6102') AND ((tbVendas.TipoNF)='1-SAIDA'));")
Conn.Execute strSQL



'Corrige desconto null
strSQL = ("UPDATE tbCompras set VlrDesconto = 0 where VlrDesconto is null;")
Conn.Execute strSQL
strSQL = ("Update tbComprasDet set VlrDesc = 0 where VlrDesc is null;")
Conn.Execute strSQL

'custo médio do churrasco
strSQL = ("update tbVendasDet set CustoMedio = 9 where IDProd = 3575;")
Conn.Execute strSQL


'ATUALIZA STATUS DO IMOBILIZADO EXAURIDO APÓS 48 MESES
'strSQL = ("UPDATE tbimobilizadocadastro AS Q1 INNER Join (SELECT * FROM overturecervej01.tbimobilizado where DataEmissao <= CURDATE() and Ciclo = 48) AS Q2 ON Q1.IDProd = Q2.IDProd SET Status = 'EXAURIDO' where Status = 'ATIVO';")
'Conn.Execute strSQL
'dá um monte de erro na efd selecionar manualmente os exauridos


'REGISTRA IMOBILIZADO 48 MESES
'DESCONSIDERA O PERÍODO DO SIMPLES NACIONAL < 01/07/17
'Quando adquirimos mercadorias para o nosso ATIVO-FIXO e que conste os impostos na respectiva nota fiscal, escrituramos no Registro de Entradas Modelo 1, 1A sem crédito do ICMS e sem crédito do IPI.
'O crédito do ICMS será escriturado no Livro de Apuração do ICMS Modelo 9 em 48 vezes.

'DoCmd.RunSQL ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.lancfiscal = 'OUTROS' WHERE (((tbCompras.DataEmissao)<#7/1/2017#));")
'DoCmd.RunSQL ("UPDATE (tbCliente INNER JOIN tbvendas ON tbCliente.IDCliente = tbvendas.Idcliente) INNER JOIN (tbCadProd INNER JOIN tbvendasdet ON (tbCadProd.IDProd = tbvendasdet.IDProd) AND (tbCadProd.IDProd = tbvendasdet.IDProd)) ON tbvendas.ID = tbvendasdet.IDVenda SET tbvendasdet.lancfiscal = 'OUTROS' WHERE (((tbvendas.DataEmissao)<#7/1/2017#));")




'Dim db As Database
'Dim rst As DAO.Recordset
'Set db = CurrentDb()

'Dim lin As Integer


'DoCmd.RunSQL ("delete * from tbimobilizado")
'DoCmd.RunSQL ("delete * from tbimobilizado_temp")
'DoCmd.RunSQL ("INSERT INTO tbImobilizado_temp ( ANO, MES, DataEmissao, Ciclo, CFOP, CFOP_ESC_DESC, LancFiscal, IDProd, DescProd, Qnt, ValorTot, BaseCalculo, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST, ChaveNF ) " & _
'"SELECT Year(DataEmissao) AS ANO, Month(dataemissao) AS MES, tbCompras.DataEmissao, 0 AS Ciclo, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, Sum(tbComprasDet.Qnt) AS SomaDeQnt, Sum(tbComprasDet.ValorTot) AS SomaDeValorTot, Sum(tbComprasDet.BaseCalculo) AS SomaDeBaseCalculo, Sum(tbComprasDet.Valor_ICMS) AS SomaDeValor_ICMS, Sum(tbComprasDet.Valor_PIS) AS SomaDeValor_PIS, Sum(tbComprasDet.Valor_Cofins) AS SomaDeValor_Cofins, Sum(tbComprasDet.Valor_IPI) AS SomaDeValor_IPI, Sum(tbComprasDet.Valor_ICMS_ST) AS SomaDeValor_ICMS_ST, tbCompras.ChaveNF " & _
'"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
'"GROUP BY Year(DataEmissao), Month(dataemissao), tbCompras.DataEmissao, 0, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, tbCompras.ChaveNF " & _
'"HAVING (((tbComprasDet.LancFiscal)='IMOBILIZADO'));")

'DoCmd.RunSQL ("INSERT INTO tbImobilizado_temp ( ANO, MES, DataEmissao, Ciclo, CFOP, CFOP_ESC_DESC, LancFiscal, IDProd, DescProd, Qnt, ValorTot, BaseCalculo, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST, ChaveNF ) " & _
'"SELECT Year(DataEmissao) AS ANO, Month(dataemissao) AS MES, tbCompras.DataEmissao, 0 AS Ciclo, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, Sum(tbComprasDet.Qnt) AS SomaDeQnt, Sum(tbComprasDet.ValorTot) AS SomaDeValorTot, Sum(tbComprasDet.BaseCalculo) AS SomaDeBaseCalculo, Sum(tbComprasDet.Valor_ICMS) AS SomaDeValor_ICMS, Sum(tbComprasDet.Valor_PIS) AS SomaDeValor_PIS, Sum(tbComprasDet.Valor_Cofins) AS SomaDeValor_Cofins, Sum(tbComprasDet.Valor_IPI) AS SomaDeValor_IPI, Sum(tbComprasDet.Valor_ICMS_ST) AS SomaDeValor_ICMS_ST, tbCompras.ChaveNF " & _
'"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
'"GROUP BY Year(DataEmissao), Month(dataemissao), tbCompras.DataEmissao, 0, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, tbCompras.ChaveNF " & _
'"HAVING (((tbComprasDet.LancFiscal)='IMOBILIZADO') AND ((Sum(tbComprasDet.Valor_ICMS))>0));")'

'Set rst = db.OpenRecordset("select * from tbImobilizado_temp")

''DoCmd.setwarnings (True)
'Dim chaveNFE As String

'Do Until rst.EOF = True
'If rst!Ciclo <> 0 Then
'Exit Do
'End If
'lin = 1
'    Do Until lin = 49

'If lin = 1 Then
'DataEmissao = rst!DataEmissao
'Else: DataEmissao = DateAdd("m", lin - 1, rst!DataEmissao)
'End If

'Ciclo = lin
'CFOP = rst!CFOP
'CFOPDesc = rst!CFOP_ESC_DESC
'LancFiscal = rst!LancFiscal
'IDProd = rst!IDProd
'DescProd = rst!DescProd
'Qnt = rst!Qnt
'ValorTot = rst!ValorTot - ((rst!ValorTot / 48) * lin)
'BaseCalculo = rst!BaseCalculo / 48
'Valor_ICMS_total = rst!Valor_ICMS
'Valor_ICMS = rst!Valor_ICMS / 48
'Valor_PIS = rst!Valor_PIS / 48
'Valor_COFINS = rst!Valor_COFINS / 48
'Valor_IPI = rst!Valor_IPI / 48
'Valor_ICMS_ST = rst!Valor_ICMS_ST / 48
'chaveNFE = rst!ChaveNF

'DoCmd.RunSQL ("INSERT INTO tbImobilizado ( ANO, MES, DataEmissao, Ciclo, CFOP, CFOP Desc,LancFiscal,                                                                                                                IDProd, DescProd, Qnt, ValorTot, BaseCalculo, Valor_ICMS_total, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST, ChaveNFe ) " & _
                                         "select " & Year(DataEmissao) & "," & Month(DataEmissao) & ",'" & Format(DataEmissao, "dd/mm/yyyy") & "'," & lin & "," & CFOP & ",'" & CFOPDesc & "','" & LancFiscal & "'," & IDProd & ",'" & DescProd & "'," & Qnt & "," & Replace(ValorTot, ",", ".") & "," & Replace(BaseCalculo, ",", ".") & "," & Replace(Valor_ICMS_total, ",", ".") & "," & Replace(Valor_ICMS, ",", ".") & "," & Replace(Valor_PIS, ",", ".") & "," & Replace(Valor_COFINS, ",", ".") & "," & Replace(Valor_IPI, ",", ".") & "," & Replace(Valor_ICMS_ST, ",", ".") & ",'" & chaveNFE & "'")

 '   lin = lin + 1

 '   Loop
    
'rst.MoveNext
'Loop

'DoCmd.RunSQL ("delete * from tbImobilizado where ciclo = '0'")
'DoCmd.RunSQL ("delete * from tbImobilizado_temp")

Call DisconnectFromDataBase

'DoCmd.setwarnings (True)

End Sub


Public Sub Imobilizacao_de_Impostos()

'REGISTRA IMOBILIZADO 48 MESES
'DESCONSIDERA O PERÍODO DO SIMPLES NACIONAL < 01/07/17
'Quando adquirimos mercadorias para o nosso ATIVO-FIXO e que conste os impostos na respectiva nota fiscal, escrituramos no Registro de Entradas Modelo 1, 1A sem crédito do ICMS e sem crédito do IPI.
'O crédito do ICMS será escriturado no Livro de Apuração do ICMS Modelo 9 em 48 vezes.

'DoCmd.RunSQL ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.lancfiscal = 'OUTROS' WHERE (((tbCompras.DataEmissao)<#7/1/2017#));")
'DoCmd.RunSQL ("UPDATE (tbCliente INNER JOIN tbvendas ON tbCliente.IDCliente = tbvendas.Idcliente) INNER JOIN (tbCadProd INNER JOIN tbvendasdet ON (tbCadProd.IDProd = tbvendasdet.IDProd) AND (tbCadProd.IDProd = tbvendasdet.IDProd)) ON tbvendas.ID = tbvendasdet.IDVenda SET tbvendasdet.lancfiscal = 'OUTROS' WHERE (((tbvendas.DataEmissao)<#7/1/2017#));")

'DoCmd.setwarnings (False)


Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

Dim lin As Integer
Call ConnectToDataBase

strSQL = ("delete from tbImobilizado where DataCompra >= '" & Format(Forms!frmXMLinput!txt_dt_corte, "yyyy-mm-dd") & "'")
Conn.Execute strSQL
strSQL = ("delete from tbImobilizado_temp")
Conn.Execute strSQL
'DoCmd.RunSQL ("INSERT INTO tbImobilizado_temp ( ANO, MES, DataEmissao, Ciclo, CFOP, CFOP_ESC_DESC, LancFiscal, IDProd, DescProd, Qnt, ValorTot, BaseCalculo, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST, ChaveNF ) " & _
'"SELECT Year(DataEmissao) AS ANO, Month(dataemissao) AS MES, tbCompras.DataEmissao, 0 AS Ciclo, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, Sum(tbComprasDet.Qnt) AS SomaDeQnt, Sum(tbComprasDet.ValorTot) AS SomaDeValorTot, Sum(tbComprasDet.BaseCalculo) AS SomaDeBaseCalculo, Sum(tbComprasDet.Valor_ICMS) AS SomaDeValor_ICMS, Sum(tbComprasDet.Valor_PIS) AS SomaDeValor_PIS, Sum(tbComprasDet.Valor_Cofins) AS SomaDeValor_Cofins, Sum(tbComprasDet.Valor_IPI) AS SomaDeValor_IPI, Sum(tbComprasDet.Valor_ICMS_ST) AS SomaDeValor_ICMS_ST, tbCompras.ChaveNF " & _
'"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
'"GROUP BY Year(DataEmissao), Month(dataemissao), tbCompras.DataEmissao, 0, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, tbCompras.ChaveNF " & _
'"HAVING (((tbComprasDet.LancFiscal)='IMOBILIZADO'));")

strSQL = ("INSERT INTO tbImobilizado_temp ( ANO, MES, DataEmissao, Ciclo, CFOP, CFOP_ESC_DESC, LancFiscal, IDProd, DescProd, Qnt, ValorTot, BaseCalculo, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST, ChaveNF ) " & _
"SELECT Year(DataEmissao) AS ANO, Month(dataemissao) AS MES, tbCompras.DataEmissao, 0 AS Ciclo, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, Sum(tbComprasDet.Qnt) AS SomaDeQnt, Sum(tbComprasDet.ValorTot) AS SomaDeValorTot, Sum(tbComprasDet.BaseCalculo) AS SomaDeBaseCalculo, Sum(tbComprasDet.Valor_ICMS) AS SomaDeValor_ICMS, Sum(tbComprasDet.Valor_PIS) AS SomaDeValor_PIS, Sum(tbComprasDet.Valor_Cofins) AS SomaDeValor_Cofins, Sum(tbComprasDet.Valor_IPI) AS SomaDeValor_IPI, Sum(tbComprasDet.Valor_ICMS_ST) AS SomaDeValor_ICMS_ST, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"GROUP BY Year(DataEmissao), Month(dataemissao), tbCompras.DataEmissao, Ciclo, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, tbCompras.ChaveNF " & _
"HAVING (((tbComprasDet.LancFiscal)='IMOBILIZADO') AND ((Sum(tbComprasDet.Valor_ICMS))>0) and DataEmissao >= '" & Format(Forms!frmXMLinput!txt_dt_corte, "yyyy-mm-dd") & "');")
Conn.Execute strSQL

Set rst = Db.OpenRecordset("select * from tbImobilizado_temp")

'DoCmd.setwarnings (False)
Dim chaveNFE As String

Do Until rst.EOF = True
If rst!Ciclo <> 0 Then
Exit Do
End If
lin = 1
    Do Until lin = 49

If lin = 1 Then
DataEmissao = rst!DataEmissao
cDataCompra = rst!DataEmissao
Else: DataEmissao = DateAdd("m", lin - 1, rst!DataEmissao)
End If

Ciclo = lin
CFOP = rst!CFOP
CFOPDesc = rst!CFOP_ESC_DESC
LancFiscal = rst!LancFiscal
IDProd = rst!IDProd
DescProd = rst!DescProd
Qnt = rst!Qnt
ValorTot = rst!ValorTot - ((rst!ValorTot / 48) * lin)
BaseCalculo = rst!BaseCalculo / 48
Valor_ICMS_total = rst!Valor_ICMS
Valor_ICMS = rst!Valor_ICMS / 48
Valor_PIS = rst!Valor_PIS / 48
Valor_Cofins = rst!Valor_Cofins / 48
Valor_IPI = rst!Valor_IPI / 48
Valor_ICMS_ST = rst!Valor_ICMS_ST / 48
chaveNFE = rst!chavenf

strSQL = ("INSERT INTO tbImobilizado ( ANO, MES, DataCompra, DataEmissao, Ciclo, CFOP, `CFOP Desc`, LancFiscal, IDProd, DescProd, Qnt, ValorTot, BaseCalculo, Valor_ICMS_total, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST, ChaveNFe ) " & _
                                         "select '" & year(DataEmissao) & "', " & month(DataEmissao) & ", '" & Format(cDataCompra, "yyyy-mm-dd") & "' , '" & Format(DataEmissao, "yyyy-mm-dd") & "' ," & lin & "," & CFOP & ",'" & CFOPDesc & "','" & LancFiscal & "'," & IDProd & ",'" & DescProd & "'," & Replace(Qnt, ",", ".") & "," & Replace(ValorTot, ",", ".") & "," & Replace(BaseCalculo, ",", ".") & "," & Replace(Valor_ICMS_total, ",", ".") & "," & Replace(Valor_ICMS, ",", ".") & "," & Replace(Valor_PIS, ",", ".") & "," & Replace(Valor_Cofins, ",", ".") & "," & Replace(Valor_IPI, ",", ".") & "," & Replace(Valor_ICMS_ST, ",", ".") & ",'" & chaveNFE & "'")
                                         
                                        

Call ConnectToDataBase
Conn.Execute strSQL
    lin = lin + 1

    Loop
    
rst.MoveNext
Loop



strSQL = ("delete from tbImobilizado where ciclo = '0'")

Call ConnectToDataBase
Conn.Execute strSQL
strSQL = ("delete from tbImobilizado_temp")
Conn.Execute strSQL


'DoCmd.setwarnings (True)


End Sub


Public Sub imobilizacao_real()
'DoCmd.setwarnings (False)
'REGISTRA IMOBILIZADO 120 MESES
'DESCONSIDERA O PERÍODO DO SIMPLES NACIONAL < 01/07/17
'NESSE CASO REFERE-SE A IMOBILIZACAO DO BEM EM SI E NÃO DOS CREDITOS DE IMPOSTOS

'DoCmd.RunSQL ("UPDATE tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor SET tbComprasDet.lancfiscal = 'OUTROS' WHERE (((tbCompras.DataEmissao)<#7/1/2017#));")
'DoCmd.RunSQL ("UPDATE (tbCliente INNER JOIN tbvendas ON tbCliente.IDCliente = tbvendas.Idcliente) INNER JOIN (tbCadProd INNER JOIN tbvendasdet ON (tbCadProd.IDProd = tbvendasdet.IDProd) AND (tbCadProd.IDProd = tbvendasdet.IDProd)) ON tbvendas.ID = tbvendasdet.IDVenda SET tbvendasdet.lancfiscal = 'OUTROS' WHERE (((tbvendas.DataEmissao)<#7/1/2017#));")




Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

Dim lin As Integer


strSQL = ("delete * from tbimobilizado_real")
Conn.Execute strSQL
strSQL = ("delete * from tbimobilizado_temp")
Conn.Execute strSQL
'DoCmd.RunSQL ("INSERT INTO tbImobilizado_temp ( ANO, MES, DataEmissao, Ciclo, CFOP, CFOP_ESC_DESC, LancFiscal, IDProd, DescProd, Qnt, ValorTot, BaseCalculo, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST, ChaveNF ) " & _
'"SELECT Year(DataEmissao) AS ANO, Month(dataemissao) AS MES, tbCompras.DataEmissao, 0 AS Ciclo, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, Sum(tbComprasDet.Qnt) AS SomaDeQnt, Sum(tbComprasDet.ValorTot) AS SomaDeValorTot, Sum(tbComprasDet.BaseCalculo) AS SomaDeBaseCalculo, Sum(tbComprasDet.Valor_ICMS) AS SomaDeValor_ICMS, Sum(tbComprasDet.Valor_PIS) AS SomaDeValor_PIS, Sum(tbComprasDet.Valor_Cofins) AS SomaDeValor_Cofins, Sum(tbComprasDet.Valor_IPI) AS SomaDeValor_IPI, Sum(tbComprasDet.Valor_ICMS_ST) AS SomaDeValor_ICMS_ST, tbCompras.ChaveNF " & _
'"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
'"GROUP BY Year(DataEmissao), Month(dataemissao), tbCompras.DataEmissao, 0, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, tbCompras.ChaveNF " & _
'"HAVING (((tbComprasDet.LancFiscal)='IMOBILIZADO'));")

strSQL = ("INSERT INTO tbImobilizado_temp ( ANO, MES, DataEmissao, Ciclo, LancFiscal, IDProd, DescProd, Qnt, ValorTot, ChaveNF ) " & _
"SELECT Year(DataEmissao) AS ANO, Month(dataemissao) AS MES, tbCompras.DataEmissao, 0 AS Ciclo, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, Sum(tbComprasDet.Qnt) AS SomaDeQnt, Sum(tbComprasDet.ValorTot) AS SomaDeValorTot, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"GROUP BY Year(DataEmissao), Month(dataemissao), tbCompras.DataEmissao, 0, tbComprasDet.LancFiscal, tbComprasDet.IDProd, tbCadProd.DescProd, tbCompras.ChaveNF " & _
"HAVING (((tbComprasDet.LancFiscal)='IMOBILIZADO'));")
Conn.Execute strSQL
'"HAVING (((tbCompras.DataEmissao)>=#1/1/2018#) AND ((tbComprasDet.LancFiscal)='IMOBILIZADO'));")

Set rst = Db.OpenRecordset("select * from tbImobilizado_temp")

'DoCmd.setwarnings (False)
Dim chaveNFE As String

Do Until rst.EOF = True
If rst!Ciclo <> 0 Then
Exit Do
End If
lin = 1
    Do Until lin = 121

If lin = 1 Then
DtCompra = rst!DataEmissao
DtImob = DateAdd("m", 1, rst!DataEmissao)
DtCiclo = DateAdd("m", 1, rst!DataEmissao)
ValorIni = rst!ValorTot
Else: DtCiclo = DateAdd("m", lin, rst!DataEmissao)
End If

Ciclo = lin
LancFiscal = rst!LancFiscal
IDProd = rst!IDProd
DescProd = rst!DescProd
Qnt = rst!Qnt
ValorTot = rst!ValorTot - ((rst!ValorTot / 120) * lin)

chaveNFE = rst!chavenf

strSQL = ("INSERT INTO tbImobilizado_Real ( ANO, MES, DtCompra, DtImob, DtCiclo, Ciclo, IDProd, DescProd, Qnt, ValorIni, ValorAtual, Depreciacao, Depreciacao_Mes, ChaveNFe ) " & _
                                         "select " & year(DtCiclo) & "," & month(DtCiclo) & ",'" & Format(DtCompra, "dd/mm/yyyy") & "','" & Format(DtImob, "dd/mm/yyyy") & "','" & Format(DtCiclo, "dd/mm/yyyy") & "'," & lin & "," & IDProd & ",'" & DescProd & "'," & Qnt & "," & Replace(ValorIni, ",", ".") & "," & Replace(ValorTot, ",", ".") & "," & Replace(ValorIni - ValorTot, ",", ".") & "," & Replace((ValorIni - ValorTot) / lin, ",", ".") & ",'" & chaveNFE & "'")
Conn.Execute strSQL

    lin = lin + 1

    Loop
    
rst.MoveNext
Loop

strSQL = ("delete * from tbImobilizado_real where ciclo = '0'")
Conn.Execute strSQL
strSQL = ("delete * from tbImobilizado_temp")
Conn.Execute strSQL


'DoCmd.setwarnings (True)

End Sub


Public Sub CalculaInventario(cIventNovo As String)
Call ConnectToDataBase

'GERAR INVENTARIO
'DoCmd.setwarnings (False)

If cIventNovo = 0 Then
MsgBox ("Lance o iventario na tabela Invetario e informe o número do ID")
GoTo fim:
Else
End If

Dim cIventAnt As Integer

Dim cDtIni As String
Dim cDtFim As String

'cDtINI = Format(cDtINI, "mm/dd/yyyy")
'cDtFIM = Format(cDtFIM, "mm/dd/yyyy")

Dim Db As Database
Set Db = CurrentDb()

Dim rsIvent As DAO.Recordset
Set rsIvent = Db.OpenRecordset("select * from tbIventario where ID = " & cIventNovo & ";")

If rsIvent.EOF = True And rsIvent.BOF = True Then
MsgBox ("Lance o iventario na tabela Invetario e informe o número do ID")
GoTo fim:
Else
End If
rsIvent.MoveFirst
cDtIni = Format(rsIvent!Data_INI, "yyyy-mm-dd")
cDtFim = Format(rsIvent!Data_FIM, "yyyy-mm-dd")

cDtIniVb = Format(rsIvent!Data_INI, "mm/dd/yyyy")
cDtFimVb = Format(rsIvent!Data_FIM, "mm/dd/yyyy")


cIventAnt = InputBox("Informe o ID iventario ANTERIOR de referencia.")

'LIMPA INVENTARIO DET ANTERIOR
strSQL = ("DELETE FROM tbIventarioDet WHERE ID_Iventario = " & cIventNovo & " ")
Conn.Execute strSQL
'LIMPA INVENTARIO DET ANTERIOR

'INSERE INVENTARIO ANTERIOR COMO NOVO
strSQL = ("INSERT INTO tbIventarioDet ( ID_Iventario, ID_Prod, Unid_Item, Qtd, Valor_Unit, Valor_Total, Valor_Item_IR ) " & _
"SELECT " & cIventNovo & " AS ID_IVENT, tbIventarioDet.ID_Prod, tbIventarioDet.Unid_Item, tbIventarioDet.Qtd, tbIventarioDet.Valor_Unit, tbIventarioDet.Valor_Total, tbIventarioDet.Valor_Item_IR " & _
"FROM tbIventarioDet " & _
"WHERE (((tbIventarioDet.ID_Iventario)=" & cIventAnt & "));")
Conn.Execute strSQL
'INSERE INVENTARIO ANTERIOR COMO NOVO

'INSERE FALTANTES
strSQL = ("INSERT INTO tbIventarioDet ( ID_Iventario, ID_Prod ) " & _
"SELECT " & cIventNovo & " AS id, QT.IDProd " & _
"FROM (SELECT Q1.IDProd, Sum(Q1.Qnt) AS Qnt, Sum(Q1.ValorTot) AS ValorTot " & _
"FROM (SELECT tbComprasDet.IDProd, Sum(tbComprasDet.Qnt) AS Qnt, Sum(tbComprasDet.ValorTot) AS ValorTot " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= '" & cDtIni & "' And (tbCompras.DataEmissao) <= '" & cDtFim & "')) " & _
"GROUP BY tbComprasDet.IDProd UNION " & _
"SELECT tbVendasDet.IDProd, Sum(tbVendasDet.Qnt *-1) AS Qnt, Sum(tbVendasDet.ValorTot * -1) AS ValorTot " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= '" & cDtIni & "' And (tbVendas.DataEmissao) <= '" & cDtFim & "')) GROUP BY tbVendasDet.IDProd UNION " & _
"SELECT tb_Registro_Envase.ID_Produto AS IDProd, Sum(tb_Registro_Envase.Qt) AS Qnt, Sum(tb_Registro_Envase.Custo_Total) AS Valor_Tot " & _
"FROM tb_Registro_Envase WHERE (((tb_Registro_Envase.Data) >= '" & cDtIni & "' And (tb_Registro_Envase.Data) <= '" & cDtFim & "')) GROUP BY tb_Registro_Envase.ID_Produto UNION " & _
"SELECT tb_Registro_Consumo.ID_Produto AS ID_Prod, Sum(tb_Registro_Consumo.Qt_Consumo) AS Qnt, Sum(tb_Registro_Consumo.Custo_Total) AS Valor_Tot " & _
"FROM tb_Registro_Consumo WHERE (((tb_Registro_Consumo.Data) >= '" & cDtIni & "' And (tb_Registro_Consumo.Data) <= '" & cDtFim & "')) GROUP BY tb_Registro_Consumo.ID_Produto)  AS Q1 GROUP BY Q1.IDProd " & _
")  AS QT LEFT JOIN tbIventarioDet ON QT.IDProd = tbIventarioDet.ID_Prod " & _
"WHERE (((tbIventarioDet.ID_Prod) Is Null));")
Conn.Execute strSQL
'INSERE FALTANTES

'PREENCHE DADOS FALTANTES
strSQL = ("UPDATE tbCadProd INNER JOIN tbIventarioDet ON tbCadProd.IDProd = tbIventarioDet.ID_Prod SET tbIventarioDet.Unid_Item = tbCadProd.Unid, tbIventarioDet.Valor_Unit = tbCadProd.CMed_Unit WHERE (((tbIventarioDet.ID_Iventario)=" & cIventNovo & "));")
Conn.Execute strSQL
'PREENCHE DADOS FALTANTES

'ZERA CAMPOS NULL
strSQL = ("UPDATE tbIventarioDet SET tbIventarioDet.Qtd = 0, tbIventarioDet.Valor_Total = 0, tbIventarioDet.Valor_Item_IR = 0 WHERE tbIventarioDet.ID_Iventario=" & cIventNovo & ";")
Conn.Execute strSQL
'ZERA CAMPOS NULL

'FAZ O LOOP
Dim rsIventDet As DAO.Recordset
Set rsIventDet = Db.OpenRecordset("select * from tbIventarioDet where ID_Iventario = " & cIventNovo & ";")

Do Until rsIventDet.EOF = True
'DEBITA AS VENDAS
Set rsVendas = Db.OpenRecordset("SELECT tbVendasDet.IDProd, Sum(tbVendasDet.Qnt) AS Qnt " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIniVb & "# And (tbVendas.DataEmissao) <= #" & cDtFimVb & "#)) " & _
"GROUP BY tbVendasDet.IDProd " & _
"HAVING (((tbVendasDet.IDProd)=" & rsIventDet!ID_Prod & "));")



If rsVendas.EOF = True And rsVendas.BOF = True Then
Else
rsIventDet.Edit
rsIventDet!Qtd = rsIventDet!Qtd - rsVendas!Qnt
rsIventDet.Update
End If
rsVendas.Close
'DEBITA CONSUMO
Set rsConsumo = Db.OpenRecordset("SELECT tb_Registro_Consumo.ID_Produto, Sum(tb_Registro_Consumo.Qt_Consumo) AS qt " & _
"FROM tb_Registro_Consumo " & _
"WHERE (((tb_Registro_Consumo.Data) >= #" & cDtIniVb & "# And (tb_Registro_Consumo.Data) <= #" & cDtFimVb & "#)) " & _
"GROUP BY tb_Registro_Consumo.ID_Produto HAVING tb_Registro_Consumo.ID_Produto=" & rsIventDet!ID_Prod & ";")

If rsConsumo.EOF = True And rsConsumo.BOF = True Then
Else
rsIventDet.Edit
rsIventDet!Qtd = rsIventDet!Qtd - rsConsumo!Qt
rsIventDet.Update
End If
rsConsumo.Close

'CREDITA AS COMPRAS
Set rsCompras = Db.OpenRecordset("SELECT tbComprasDet.IDProd, Sum(tbComprasDet.Qnt) AS Qnt " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= #" & cDtIniVb & "# And (tbCompras.DataEmissao) <= #" & cDtFimVb & "#)) " & _
"GROUP BY tbComprasDet.IDProd " & _
"HAVING (((tbComprasDet.IDProd)= " & rsIventDet!ID_Prod & "));")

If rsCompras.EOF = True And rsCompras.BOF = True Then
Else
rsIventDet.Edit
rsIventDet!Qtd = rsIventDet!Qtd + rsCompras!Qnt
rsIventDet.Update
End If
rsCompras.Close
'CREDITA ENVASE
Set rsEnvase = Db.OpenRecordset("SELECT tb_Registro_Envase.ID_Produto, Sum(tb_Registro_Envase.Qt) AS Qt " & _
"FROM tb_Registro_Envase " & _
"WHERE (((tb_Registro_Envase.Data) >= #" & cDtIniVb & "# And (tb_Registro_Envase.Data) <= #" & cDtFimVb & "#)) " & _
"GROUP BY tb_Registro_Envase.ID_Produto " & _
"HAVING (((tb_Registro_Envase.ID_Produto)=" & rsIventDet!ID_Prod & "));")

If rsEnvase.EOF = True And rsEnvase.BOF = True Then
Else
rsIventDet.Edit
rsIventDet!Qtd = rsIventDet!Qtd + rsEnvase!Qt
rsIventDet.Update
End If
rsEnvase.Close


rsIventDet.MoveNext
Loop

'AJUSTE DE NEGATIVOS

Call ConnectToDataBase
strSQL = ("UPDATE tbIventarioDet set qtd = 0 where qtd <0 and ID_Iventario = " & cIventNovo & "")
Conn.Execute strSQL
'AJUSTE DE NEGATIVOS

'CALCULA VALOR TOTAL
strSQL = ("UPDATE tbIventarioDet set Valor_Total = Valor_unit * Qtd, Valor_Item_IR = Valor_unit * qtd where ID_Iventario = " & cIventNovo & "")
Conn.Execute strSQL
'CALCULA VALOR TOTAL

'DELETA ENERGIA ELETRICA
strSQL = ("DELETE FROM tbIventarioDet where ID_PROD = 2800 and Id_Iventario = " & cIventNovo & "")
Conn.Execute strSQL
'DELETA ENERGIA ELÉTRICA

'DELETA IMOBILIZADO
strSQL = ("DELETE tbIventarioDet " & _
"FROM tbCadProd INNER JOIN tbIventarioDet ON tbCadProd.IDProd = tbIventarioDet.ID_Prod " & _
"WHERE tbCadProd.IMOBILIZADO='SIM' AND tbIventarioDet.ID_Iventario= " & cIventNovo & ";")
Conn.Execute strSQL

'DELETA IMOBILIZADO

'ATUALIZAR VALOR TOTAL DO ESTOQUE
Set rsEnvase = Db.OpenRecordset("SELECT Sum(tbIventarioDet.Valor_Total) AS Valor_Total FROM tbIventarioDet " & _
"WHERE (((tbIventarioDet.ID_Iventario)=" & cIventNovo & "));")
Dim cEstoquetotal As Variant
cEstoquetotal = CDec(rsEnvase!Valor_Total)

strSQL = ("update tbIventario set tbIventario.ValorTotalEstoque = " & Replace(cEstoquetotal, ",", ".") & " where ID = " & cIventNovo & ";")
Conn.Execute strSQL
'ATUALIZAR VALOR TOTAL DO ESTOQUE
MsgBox ("OK")

fim:
'DoCmd.setwarnings (True)
Call DisconnectFromDataBase

End Sub


Private Sub Calc_CustoProducao()
'ESSE CÓDIGO VAI RODAR O CUSTO DE PRODUÇÃO APENAS DE PRODUTOS PRODUZIDOS FAZENDO O RATEIO POR CUSTO MENSAL
'DE INSUMOS, ENERGIA ELÉTRICA E EMBALAGEM
'NÃO TEM MÃO DE OBRA AINDA

Call ConnectToDataBase


strSQL = ("DELETE FROM tb_Registro_CustoProducao_Total;")
Conn.Execute strSQL

strSQL = ("DELETE FROM tb_Registro_CustoProducao;")
Conn.Execute strSQL

'CRIA ANO MES
strSQL = ("insert tb_Registro_CustoProducao_Total (ANO, MES) select * from " & _
"(select ANO, MES from " & _
"(SELECT ANO, MES, q1.ID, q2.IDProd FROM tbCompras as q1 " & _
"inner join tbComprasDet as q2 " & _
"on q1.ID = q2.IDCompra " & _
"inner join tbCadProd as q3 " & _
"on q2.IDProd = q3.IDProd " & _
") AS QDet group by Ano, Mes) as QTot;")
Conn.Execute strSQL


strSQL = ("insert tb_Registro_CustoProducao (ANO, MES, IDProd) select * from " & _
"(select ANO, MES, IDProd from " & _
"(SELECT ANO, MES, q1.ID, q2.IDProd, q3.DescProd FROM tbVendas as q1 " & _
"inner join tbVendasDet as q2 " & _
"on q1.ID = q2.IDVenda " & _
"inner join tbCadProd as q3 " & _
"on q2.IDProd = q3.IDProd where q3.PROD_FINAL = 'SIM' " & _
") AS QDet group by Ano, Mes, IDProd) as QTot;")
Conn.Execute strSQL


strSQL = ("update tb_Registro_CustoProducao_Total set Qt_Vend = 0, Custo_Insumos = 0, Custo_Embalagem = 0, Custo_Energia = 0, CUsto_Total = 0;")
Conn.Execute strSQL

strSQL = ("update tb_Registro_CustoProducao set Qt_Vend = 0, Custo_Insumos = 0, Custo_Embalagem = 0, Custo_Energia = 0, CUsto_Total = 0, Custo_Unit = 0;")
Conn.Execute strSQL

strSQL = ("update tb_Registro_CustoProducao_Total as q0 inner join " & _
"(select ANO, MES, sum(Custo) as Custo from " & _
"(SELECT ANO, MES, q1.ID, q2.IDProd, q3.DescProd, q2.ValorTot-q2.Valor_ICMS -q2.Valor_IPI - q2.Valor_Pis - q2.Valor_Cofins as Custo FROM tbCompras as q1 " & _
"inner join tbComprasDet as q2 " & _
"on q1.ID = q2.IDCompra " & _
"inner join tbCadProd as q3 " & _
"on q2.IDProd = q3.IDProd " & _
"where q3.MAT_PRIMA = 'SIM' and IdFornecedor <> 1131) AS QDet group by Ano, Mes) as q1 set q0.Custo_Insumos = q1.Custo WHERE q0.ANO = q1.ANO AND q0.MES = q1.MES;")
Conn.Execute strSQL


strSQL = ("update tb_Registro_CustoProducao_Total as q0 inner join " & _
"(select ANO, MES, sum(Custo) as Custo from " & _
"(SELECT ANO, MES, q1.ID, q2.IDProd, q3.DescProd, q2.ValorTot-q2.Valor_ICMS -q2.Valor_IPI - q2.Valor_Pis - q2.Valor_Cofins as Custo FROM tbCompras as q1 " & _
"inner join tbComprasDet as q2 " & _
"on q1.ID = q2.IDCompra " & _
"inner join tbCadProd as q3 " & _
"on q2.IDProd = q3.IDProd " & _
"where q3.EMBALAGEM = 'SIM') AS QDet group by Ano, Mes) as q1 set q0.Custo_Embalagem = q1.Custo WHERE q0.ANO = q1.ANO AND q0.MES = q1.MES;")
Conn.Execute strSQL


'energia, aqui utiliza a base de calculo para considerar o valor total do kwh por causa do crédito da sunmobi
strSQL = ("update tb_Registro_CustoProducao_Total as q0 inner join " & _
"(select ANO, MES, sum(Custo) as Custo from " & _
"(SELECT ANO, MES, q1.ID, q2.IDProd, q3.DescProd, q2.BaseCalculo-q2.Valor_ICMS -q2.Valor_IPI - q2.Valor_Pis - q2.Valor_Cofins as Custo FROM tbCompras as q1 " & _
"inner join tbComprasDet as q2 " & _
"on q1.ID = q2.IDCompra " & _
"inner join tbCadProd as q3 " & _
"on q2.IDProd = q3.IDProd " & _
"where q3.MAT_PRIMA = 'SIM' and IdFornecedor = 1131) AS QDet group by Ano, Mes) as q1 set q0.Custo_Energia = q1.Custo WHERE q0.ANO = q1.ANO AND q0.MES = q1.MES;")
Conn.Execute strSQL





strSQL = ("update tb_Registro_CustoProducao_Total set Custo_Total = Custo_Insumos+Custo_Embalagem+Custo_Energia;")
Conn.Execute strSQL

strSQL = ("update tb_Registro_CustoProducao_Total as q0 inner join (select ANO, MES,            SUM(Qnt*Litros) as Qt from tbVendasDet as q0 inner join tbVendas as q3 on q0.IdVenda = q3.ID inner join tbCadProd as q1 on q0.IDProd = q1.IDProd where PROD_FINAL = 'SIM' group by Ano, Mes)         as q1 set q0.QT_Vend = q1.Qt where q0.Ano = q1.Ano and q0.Mes = q1.Mes ;")
Conn.Execute strSQL

strSQL = ("update tb_Registro_CustoProducao as q0       inner join (select ANO, MES, q0.IDProd as PROD, SUM(Qnt*Litros) as Qt from tbVendasDet as q0 inner join tbVendas as q3 on q0.IdVenda = q3.ID inner join tbCadProd as q1 on q0.IDProd = q1.IDProd where PROD_FINAL = 'SIM' group by Ano, Mes, PROD) as q1 set q0.QT_Vend = q1.Qt where q0.Ano = q1.Ano and q0.Mes = q1.Mes and q0.IDProd = q1.PROD;")
Conn.Execute strSQL

strSQL = ("UPDATE tb_Registro_CustoProducao as q0 inner join tb_Registro_CustoProducao_Total as q1 " & _
"on q1.ANO = q0.ANO and q1.MES = q0.MES " & _
"SET q0.Custo_Insumos = round((q0.QT_Vend / q1.QT_Vend) * q1.Custo_Insumos,2) ;")
Conn.Execute strSQL


strSQL = ("UPDATE tb_Registro_CustoProducao as q0 inner join tb_Registro_CustoProducao_Total as q1 " & _
"on q1.ANO = q0.ANO and q1.MES = q0.MES " & _
"SET q0.Custo_Embalagem = round((q0.QT_Vend / q1.QT_Vend) * q1.Custo_Embalagem,2) ;")
Conn.Execute strSQL


strSQL = ("UPDATE tb_Registro_CustoProducao as q0 inner join tb_Registro_CustoProducao_Total as q1 " & _
"on q1.ANO = q0.ANO and q1.MES = q0.MES " & _
"SET q0.Custo_Energia = round((q0.QT_Vend / q1.QT_Vend) * q1.Custo_Energia,2) ;")
Conn.Execute strSQL


strSQL = ("update tb_Registro_CustoProducao set Custo_Total = round(Custo_Insumos + Custo_Embalagem + Custo_Energia,2);")
Conn.Execute strSQL

strSQL = ("update tb_Registro_CustoProducao set Custo_Unit = round(Custo_Total/QT_Vend,2);")
Conn.Execute strSQL

strSQL = ("update tb_Registro_CustoProducao as q0 inner join tbCadProd as q1 on q0.IDProd = q1.IDProd set q0.Litros = q1.Litros;")
Conn.Execute strSQL

strSQL = ("update tb_Registro_CustoProducao  set Custo_Unit = round((Custo_Total / QT_Vend) * Litros,2);")
Conn.Execute strSQL


'buscar custo médio do produto de revenda

'buscar custo médio do produto de revenda



'SELECT * FROM tb_Registro_CustoProducao_Total WHERE ANO = 2020 AND MES = 4;
'SELECT q1.*, q2.DescProd FROM tb_Registro_CustoProducao q1
'inner join tbCadProd q2 on q1.IDProd = q2.IDProd  WHERE ANO = 2019 AND MES = 12;

strSQL = ("update tbVendasDet q0 inner join tbVendas q2 on q0.IDVenda = q2.ID inner join tb_Registro_CustoProducao q1 on q0.IDProd = q1.IDProd and q2.ANO = q1.ANO and q2.MES = q1.MES set q0.CustoMedio = q1.Custo_Unit * q0.Qnt;")
Conn.Execute strSQL

strSQL = ("update tbCadProd q0 inner join (SELECT q0.* FROM tb_Registro_CustoProducao q0 inner join (SELECT ANO, MES FROM (select distinct ANO, MES from tb_Registro_CustoProducao order by ano desc, mes desc) AS Q1 LIMIT 1) q1 on q0.ANO = q1.ANO and q1.MES = q0.MES) q1 on q0.IDProd = q1.IDProd set q0.CMed_Unit = q1.Custo_Unit, q0.CMED_Tot = round(q1.Custo_Unit * Estoque,2);")
Conn.Execute strSQL

Call DisconnectFromDataBase


End Sub







