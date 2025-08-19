Attribute VB_Name = "modTriggers"
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


Public Sub Atualiza_CFOP_DESC()

Call ConnectToDataBase

'DoCmd.setwarnings (False)
'CFOP COMPRA
strSQL = ("UPDATE tbCFOP_3 INNER JOIN tbComprasDet ON tbCFOP_3.NIV3 = tbComprasDet.CFOP SET tbComprasDet.CFOP_DESC = tbCFOP_3.DESC WHERE CFOP_DESC IS NULL;")
Conn.Execute strSQL
'CFOP VENDA
strSQL = ("UPDATE tbCFOP_3 INNER JOIN tbVendasDet ON tbCFOP_3.NIV3 = tbVendasDet.CFOP SET tbVendasDet.CFOP_DESC = tbCFOP_3.DESC WHERE CFOP_DESC IS NULL;")
Conn.Execute strSQL
'CST COMPRA
strSQL = ("UPDATE tbCST INNER JOIN tbComprasDet ON tbCST.CST = tbComprasDet.CST SET tbComprasDet.CST_DESC = tbCST.Desc where tbComprasDet.CST_DESC is NULL;")
Conn.Execute strSQL
'CST VENDA
strSQL = ("UPDATE tbCST INNER JOIN tbVendasDet ON tbCST.CST = tbVendasDet.CST SET tbVendasDet.CST_DESC = tbCST.Desc WHERE (((tbVendasDet.CST_DESC) Is Null));")
Conn.Execute strSQL
'ORIGEM COMPRA

strSQL = ("UPDATE tbOrigem INNER JOIN tbComprasDet ON tbOrigem.CD_ORIGEM = tbComprasDet.Origem SET tbComprasDet.Origem = tbOrigem.Desc;")
Conn.Execute strSQL
'ORIGEM VENDA
strSQL = ("UPDATE tbOrigem INNER JOIN tbVendasDet ON tbOrigem.CD_ORIGEM = tbVendasDet.Origem SET tbVendasDet.Origem = tbOrigem.Desc;")
Conn.Execute strSQL

'ATUALIZA COD MUNICIPIO CLIENTES
strSQL = ("UPDATE tbCliente INNER JOIN tbMunicipioIBGE ON tbCliente.Municipio = tbMunicipioIBGE.MUNICIPIO SET tbCliente.cod_Municipio = tbMunicipioIBGE.COD_MUNICIPIO WHERE (((tbCliente.cod_Municipio) Is Null));")
Conn.Execute strSQL
strSQL = ("UPDATE tbCliente INNER JOIN tbMunicipioIBGE ON tbCliente.Municipio = tbMunicipioIBGE.MUNICIPIO_ACENTO SET tbCliente.cod_Municipio = tbMunicipioIBGE.COD_MUNICIPIO WHERE (((tbCliente.cod_Municipio) Is Null));")
Conn.Execute strSQL

'ATUALIZA COD MUNICIPIO FORNECEDORES
strSQL = ("UPDATE tbFornecedor INNER JOIN tbMunicipioIBGE ON tbFornecedor.Municipio = tbMunicipioIBGE.MUNICIPIO SET tbFornecedor.cod_Municipio = tbMunicipioIBGE.COD_MUNICIPIO WHERE (((tbFornecedor.cod_Municipio) Is Null));")
Conn.Execute strSQL
strSQL = ("UPDATE tbFornecedor INNER JOIN tbMunicipioIBGE ON tbFornecedor.Municipio = tbMunicipioIBGE.MUNICIPIO_ACENTO SET tbFornecedor.cod_Municipio = tbMunicipioIBGE.COD_MUNICIPIO WHERE (((tbFornecedor.cod_Municipio) Is Null));")
Conn.Execute strSQL

'ATUALIZA NF STATUS ATIVO NOVAS
strSQL = ("UPDATE tbVendas set Status = 'ATIVO' WHERE Status = NULL;")
Conn.Execute strSQL

'Genero Fiscal materia prima
'11  Produtos da indústria de moagem; malte; amidos e féculas; inulina; glúten de trigo
'DoCmd.RunSQL ("UPDATE tbCadProd SET tbCadProd.Genero_Fiscal = '11' WHERE (((tbCadProd.DescProd) Like '*malte*'));")
'14  Matérias para entrançar e outros produtos de origem vegetal, não especificadas nem compreendidas em outros Capítulos da NCM
'DoCmd.RunSQL ("UPDATE tbCadProd SET tbCadProd.Genero_Fiscal = '14' WHERE (((tbCadProd.DescProd) Like '*lupulo*'));")

'ATUALIZA BASE DE CALCULO DO PIS COFINS SE FOR NULO
strSQL = ("UPDATE tbComprasDet set BaseCalc_PisCofins = BaseCalculo where BaseCalc_PisCofins is null")
Conn.Execute strSQL

strSQL = ("UPDATE tbVendasDet set BaseCalc_PisCofins = BaseCalculo where BaseCalc_PisCofins is null")
Conn.Execute strSQL


'INSERE TABELA IMOBILIZADO
strSQL = ("INSERT INTO tbImobilizadoCadastro ( IDProd, Descricao, Nr_Parcelas, ID_Conta ) " & _
"SELECT tbCadProd.IDProd, tbCadProd.DescProd, 48 AS Parc, 1 as ID " & _
"FROM tbCadProd LEFT JOIN tbImobilizadoCadastro ON tbCadProd.IDProd = tbImobilizadoCadastro.IDProd " & _
"WHERE (((tbImobilizadoCadastro.IDProd) Is Null) AND ((tbCadProd.IMOBILIZADO)='SIM'));")
Conn.Execute strSQL

'ATUALIZA ANO E MES DAS TABELAS
'compras
strSQL = ("UPDATE tbCompras SET tbCompras.ANO = Year(tbCompras.DataEmissao), tbCompras.MES = Month(tbCompras.DataEmissao) WHERE (((tbCompras.MES) Is Null) AND ((tbCompras.ANO) Is Null));")
Conn.Execute strSQL

'energia
strSQL = ("UPDATE tbEnergia SET tbEnergia.ANO = Year(tbEnergia.DataNota), tbEnergia.MES = Month(tbEnergia.DataNota) WHERE (((tbEnergia.MES) Is Null) AND ((tbEnergia.ANO) Is Null));")
Conn.Execute strSQL
'transportes
strSQL = ("UPDATE tbTransportes SET tbTransportes.ANO = Year(tbTransportes.DataEmissao), tbTransportes.MES = Month(tbTransportes.DataEmissao) WHERE (((tbTransportes.ANO) Is Null) AND ((tbTransportes.MES) Is Null));")
Conn.Execute strSQL
'vendas
strSQL = ("UPDATE tbVendas SET tbVendas.ANO = Year(tbVendas.DataEmissao), tbVendas.MES = Month(tbVendas.DataEmissao) WHERE (((tbVendas.ANO) Is Null) AND ((tbVendas.MES) Is Null));")
Conn.Execute strSQL



'ATUALIZA TOTAL IVENTARIO

Dim rsInv As DAO.Recordset
Dim Db As Database
Set Db = CurrentDb()

Set rsInv = Db.OpenRecordset("SELECT tbIventarioDet.ID_Iventario, Sum(tbIventarioDet.Valor_Total) AS ValorTotal FROM tbIventarioDet GROUP BY tbIventarioDet.ID_Iventario;")

Do Until rsInv.EOF
strSQL = ("UPDATE tbIventario SET tbIventario.ValorTotalEstoque = '" & Replace(rsInv!ValorTotal, ",", ".") & "' WHERE (((tbIventario.ID)= " & rsInv!ID_Iventario & " ));")
Conn.Execute strSQL
rsInv.MoveNext
Loop


Call DisconnectFromDataBase

'DoCmd.setwarnings (True)

End Sub


Public Sub CalcularSaldosImpostos()

Call ConnectToDataBase


Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

Dim lin As Integer
Dim credtransp As Double

Dim cICMS_Ciap As Double

'DoCmd.setwarnings (True)


strSQL = ("delete from tbResumo_ICMS")
Conn.Execute strSQL
strSQL = ("delete from tbResumo_ICMS_ST")
Conn.Execute strSQL
strSQL = ("delete from tbResumo_IPI")
Conn.Execute strSQL
strSQL = ("delete from tbResumo_PIS")
Conn.Execute strSQL
strSQL = ("delete from tbResumo_Cofins")
Conn.Execute strSQL
strSQL = ("delete from tbResumo_ICMS_ST")
Conn.Execute strSQL
strSQL = ("delete from tbResumo_IRPJ_CSLL")
Conn.Execute strSQL
strSQL = ("delete from tbResumoVendas")


'ICMS - DÁ CREDITO DE IMOBILIZADO
'DoCmd.RunSQL ("INSERT INTO tbResumo_ICMS ( ANO, MES, CRED, DEB, Saldo_Mes, CIAP, CIAP_EM_NF) " & _
'"SELECT tbCalendario.ANO, Int(tbCalendario.Mes) AS MESs, Sum(Nz(cstAcumCompras.Valor_ICMS)+Nz(cstAcumTransportes.Valor_ICMS)) AS CRED, Sum(Nz(cstAcumVendas.Valor_ICMS,0)) AS DEB, Sum((Nz(cstAcumCompras.Valor_ICMS)+Nz(cstAcumTransportes.Valor_ICMS))-(Nz(cstAcumVendas.Valor_ICMS,0))) AS Saldo_Mes, Sum(Nz(cstAcumImob.Valor_ICMS)) AS Saldo_Ciap " & _
'"FROM (((tbCalendario LEFT JOIN cstAcumCompras ON (tbCalendario.MES = cstAcumCompras.MES) AND (tbCalendario.ANO = cstAcumCompras.ANO)) LEFT JOIN cstAcumTransportes ON (tbCalendario.MES = cstAcumTransportes.MES) AND (tbCalendario.ANO = cstAcumTransportes.ANO)) LEFT JOIN cstAcumVendas ON (tbCalendario.MES = cstAcumVendas.MES) AND (tbCalendario.ANO = cstAcumVendas.ANO)) LEFT JOIN cstAcumImob ON (tbCalendario.MES = cstAcumImob.MESs) AND (tbCalendario.ANO = cstAcumImob.ANOs) " & _
'"GROUP BY tbCalendario.ANO, Int(tbCalendario.Mes) " & _
'"HAVING (((tbCalendario.ANO)='2017') AND ((Int(tbCalendario.Mes))>=7)) OR (((tbCalendario.ANO)>='2018')) " & _
'"ORDER BY tbCalendario.ANO, Int(tbCalendario.Mes);")
'"Valor_ICMS_Compras - DEB +Valor_ICMS_Transportes as Saldo_Mes, Saldo_Ciap, CASE WHEN Saldo_Ciap_NF = 0 THEN CASE WHEN ANO <=2019 THEN Saldo_Ciap ELSE Saldo_Ciap_NF END else Saldo_Ciap_NF END AS Saldo_Ciap_NF  FROM "
strSQL = ("INSERT INTO tbResumo_ICMS ( ANO, MES, CRED, DEB, Saldo_Mes, CIAP, CIAP_EM_NF) " & _
"SELECT ANO, MESs, " & _
"Valor_ICMS_Compras + Valor_ICMS_Transportes+ValorTotCIAP as CRED, DEB, " & _
"Valor_ICMS_Compras - DEB +Valor_ICMS_Transportes+ValorTotCIAP as Saldo_Mes, Saldo_Ciap, Saldo_Ciap_NF FROM " & _
"(SELECT tbCalendario.ANO, cast(tbCalendario.Mes as unsigned) AS MESs, " & _
"COALESCE(sum(cstacumcompras_credito.Valor_ICMS),0) as Valor_ICMS_Compras, " & _
"COALESCE(sum(cstAcumTransportes.Valor_ICMS),0) as Valor_ICMS_Transportes, " & _
"COALESCE(sum(cstAcumCiapNF.Valor_ICMS),0) AS ValorTotCIAP, " & _
"COALESCE(sum(cstAcumImob.Valor_ICMS),0) as Saldo_Ciap, " & _
"COALESCE(sum(cstAcumVendas.Valor_ICMS),0) AS DEB, " & _
"COALESCE(Sum(cstAcumCiapNF.Valor_ICMS), 0) As Saldo_Ciap_NF " & _
"FROM ((((tbCalendario LEFT JOIN cstAcumCompras_credito ON (tbCalendario.ANO = cstAcumCompras_credito.ANO) AND (tbCalendario.MES = cstAcumCompras_credito.MES)) LEFT JOIN cstAcumTransportes ON (tbCalendario.ANO = cstAcumTransportes.ANO) AND (tbCalendario.MES = cstAcumTransportes.MES)) LEFT JOIN cstAcumVendas ON (tbCalendario.ANO = cstAcumVendas.ANO) AND (tbCalendario.MES = cstAcumVendas.MES)) LEFT JOIN cstAcumImob ON (tbCalendario.ANO = cstAcumImob.ANOs) AND (tbCalendario.MES = cstAcumImob.MESs)) LEFT JOIN cstAcumCiapNF ON (tbCalendario.MES = cstAcumCiapNF.MES) AND (tbCalendario.ANO = cstAcumCiapNF.ANO) " & _
"GROUP BY tbCalendario.ANO, cast(tbCalendario.Mes as unsigned) " & _
"HAVING (tbCalendario.ANO='2017' AND MESs >=7) OR (tbCalendario.ANO>='2018') " & _
"ORDER BY tbCalendario.ANO, MESs " & _
") AS Q1;")
Conn.Execute strSQL



lin = 1
credtransp = 0
Set rst = Db.OpenRecordset("tbResumo_ICMS")

Do Until rst.EOF
    If lin = 1 Then
    rst.Edit
        rst!CRED_MES_ANT = 0
        'rst!CRED_TRANSPORTAR = rst!Saldo_Mes + rst!CIAP_EM_NF
        rst!CRED_TRANSPORTAR = rst!Saldo_Mes
        rst!SALDO = rst!Saldo_Mes
    rst.Update
    
    Else
        rst.MovePrevious
        credtransp = rst!CRED_TRANSPORTAR
        rst.MoveNext
        rst.Edit
        rst!CRED_MES_ANT = credtransp
            If rst!Saldo_Mes + credtransp <= 0 Then
            'rst!CRED_TRANSPORTAR = 0 + rst!CIAP_EM_NF
            rst!CRED_TRANSPORTAR = 0
            Else
            'rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp + rst!CIAP_EM_NF
            rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp
            End If
            rst!SALDO = rst!Saldo_Mes + credtransp
        rst.Update
    End If
    
         
lin = lin + 1
rst.MoveNext

Loop
rst.Close


'ICMS_ST
strSQL = ("INSERT INTO tbResumo_ICMS_ST ( ANO, MES, CRED, DEB, Saldo_Mes) " & _
"SELECT ANO, MESs, " & _
"CRED, DEB, " & _
"CRED - DEB as Saldo_Mes FROM " & _
"(SELECT tbCalendario.ANO, cast(tbCalendario.Mes as unsigned) AS MESs, " & _
"COALESCE(sum(cstAcumCompras.Valor_ICMS_ST),0) as CRED, " & _
"COALESCE(sum(cstAcumVendas.Valor_ICMS_ST),0) AS DEB " & _
"FROM ((((tbCalendario LEFT JOIN cstAcumCompras ON (tbCalendario.ANO = cstAcumCompras.ANO) AND (tbCalendario.MES = cstAcumCompras.MES)) LEFT JOIN cstAcumTransportes ON (tbCalendario.ANO = cstAcumTransportes.ANO) AND (tbCalendario.MES = cstAcumTransportes.MES)) LEFT JOIN cstAcumVendas ON (tbCalendario.ANO = cstAcumVendas.ANO) AND (tbCalendario.MES = cstAcumVendas.MES)) LEFT JOIN cstAcumImob ON (tbCalendario.ANO = cstAcumImob.ANOs) AND (tbCalendario.MES = cstAcumImob.MESs)) LEFT JOIN cstAcumCiapNF ON (tbCalendario.MES = cstAcumCiapNF.MES) AND (tbCalendario.ANO = cstAcumCiapNF.ANO) " & _
"GROUP BY tbCalendario.ANO, cast(tbCalendario.Mes as unsigned) " & _
"HAVING (tbCalendario.ANO='2017' AND MESs >=7) OR (tbCalendario.ANO>='2018') " & _
"ORDER BY tbCalendario.ANO, MESs " & _
") AS Q1;")
Conn.Execute strSQL



lin = 1
credtransp = 0
Set rst = Db.OpenRecordset("tbResumo_ICMS_ST")

Do Until rst.EOF
    If lin = 1 Then
    rst.Edit
        rst!CRED_MES_ANT = 0
        'rst!CRED_TRANSPORTAR = rst!Saldo_Mes + rst!CIAP_EM_NF
        rst!CRED_TRANSPORTAR = rst!Saldo_Mes
        rst!SALDO = rst!Saldo_Mes
    rst.Update
    
    Else
        rst.MovePrevious
        credtransp = rst!CRED_TRANSPORTAR
        rst.MoveNext
        rst.Edit
        rst!CRED_MES_ANT = credtransp
            If rst!Saldo_Mes + credtransp <= 0 Then
            'rst!CRED_TRANSPORTAR = 0 + rst!CIAP_EM_NF
            rst!CRED_TRANSPORTAR = 0
            Else
            'rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp + rst!CIAP_EM_NF
            rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp
            End If
            rst!SALDO = rst!Saldo_Mes + credtransp
        rst.Update
    End If
    
         
lin = lin + 1
rst.MoveNext

Loop
rst.Close


'IPI - IPI NÃO DÁ CREDITO DE IMOBILIZADO
strSQL = ("INSERT INTO tbResumo_IPI ( ANO, MES, CRED, DEB, Saldo_Mes ) " & _
"SELECT tbCalendario.ANO, CAST(tbCalendario.mes AS UNSIGNED) AS MESs, COALESCE(Sum(cstAcumCompras.Valor_IPI),0) AS CRED_IPI, " & _
"COALESCE(Sum(cstAcumVendas.Valor_IPI),0) AS DEB_IPI, COALESCE(Sum(cstAcumCompras.Valor_IPI),0)-COALESCE(Sum(cstAcumVendas.Valor_IPI),0) AS Saldo_Mes_IPI " & _
"FROM ((tbCalendario LEFT JOIN cstAcumCompras ON (tbCalendario.MES = cstAcumCompras.MES) AND (tbCalendario.ANO = cstAcumCompras.ANO)) LEFT JOIN cstAcumTransportes ON (tbCalendario.MES = cstAcumTransportes.MES) AND (tbCalendario.ANO = cstAcumTransportes.ANO)) LEFT JOIN cstAcumVendas ON (tbCalendario.MES = cstAcumVendas.MES) AND (tbCalendario.ANO = cstAcumVendas.ANO) " & _
"GROUP BY tbCalendario.ANO, MESs " & _
"HAVING (tbCalendario.ANO='2017' AND MESs>=7) OR (tbCalendario.ANO>='2018') " & _
"ORDER BY tbCalendario.ANO, MESs ;")
Conn.Execute strSQL

lin = 1
credtransp = 0
Set rst = Db.OpenRecordset("tbResumo_IPI")

Do Until rst.EOF
    If lin = 1 Then
    
    rst.Edit
        rst!CRED_MES_ANT = 0
        rst!CRED_TRANSPORTAR = rst!Saldo_Mes
        rst!SALDO = rst!Saldo_Mes
    rst.Update
    
    Else
        rst.MovePrevious
        credtransp = rst!CRED_TRANSPORTAR
        rst.MoveNext
        rst.Edit
        rst!CRED_MES_ANT = credtransp
            If rst!Saldo_Mes + credtransp <= 0 Then
            rst!CRED_TRANSPORTAR = 0
            Else
            rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp
            End If
            rst!SALDO = rst!Saldo_Mes + credtransp
        rst.Update
    End If
    
lin = lin + 1
rst.MoveNext
Loop
rst.Close

'PIS

strSQL = ("INSERT INTO tbResumo_PIS ( ANO, MES, CRED, DEB, Saldo_Mes ) " & _
"SELECT tbCalendario.ANO, CAST(tbCalendario.MES AS UNSIGNED) AS MESs, coalesce(Sum(cstAcumCompras_credito.Valor_PIS),0)+coalesce(sum(cstAcumTransportes.Valor_PIS),0) AS CRED_PIS, coalesce(Sum(cstAcumVendas.Valor_PIS),0) AS DEB_PIS, coalesce(Sum(cstAcumCompras_Credito.Valor_PIS),0)+coalesce(SUM(cstAcumTransportes.Valor_PIS),0)-coalesce(Sum(cstAcumVendas.Valor_PIS),0) AS Saldo_Mes_PIS " & _
"FROM ((tbCalendario LEFT JOIN cstAcumCompras_Credito ON (tbCalendario.ANO = cstAcumCompras_credito.ANO) AND (tbCalendario.MES = cstAcumCompras_credito.MES)) LEFT JOIN cstAcumTransportes ON (tbCalendario.ANO = cstAcumTransportes.ANO) AND (tbCalendario.MES = cstAcumTransportes.MES)) LEFT JOIN cstAcumVendas ON (tbCalendario.ANO = cstAcumVendas.ANO) AND (tbCalendario.MES = cstAcumVendas.MES) " & _
"GROUP BY tbCalendario.ANO, MESs " & _
"HAVING (tbCalendario.ANO='2017' AND MESs>=7) OR (tbCalendario.ANO>='2018') " & _
"ORDER BY tbCalendario.ANO, MESs ;")
Conn.Execute strSQL


lin = 1
credtransp = 0
Set rst = Db.OpenRecordset("tbResumo_PIS")

Do Until rst.EOF
    If lin = 1 Then
    
    rst.Edit
        rst!CRED_MES_ANT = 0
        rst!CRED_TRANSPORTAR = rst!Saldo_Mes
        rst!SALDO = rst!Saldo_Mes
    rst.Update
    
    Else
        rst.MovePrevious
        credtransp = rst!CRED_TRANSPORTAR
        rst.MoveNext
        rst.Edit
        rst!CRED_MES_ANT = credtransp
            If rst!Saldo_Mes + credtransp <= 0 Then
            rst!CRED_TRANSPORTAR = 0
            Else
            rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp
            End If
            rst!SALDO = rst!Saldo_Mes + credtransp
        rst.Update
    End If
    
lin = lin + 1
rst.MoveNext
Loop
rst.Close

'COFINS
strSQL = ("INSERT INTO tbResumo_Cofins ( ANO, MES, CRED, DEB, Saldo_Mes ) " & _
"SELECT tbCalendario.ANO, CAST(tbCalendario.MES AS UNSIGNED) AS MESs, coalesce(Sum(cstAcumCompras_credito.Valor_COFINS),0)+coalesce(sum(cstAcumTransportes.Valor_COFINS),0) AS CRED_COFINS, coalesce(Sum(cstAcumVendas.Valor_Cofins),0) AS DEB_COFINS, coalesce(Sum(cstAcumCompras_credito.Valor_COFINS),0)+coalesce(SUM(cstAcumTransportes.Valor_COFINS),0)-coalesce(Sum(cstAcumVendas.Valor_Cofins),0) AS Saldo_Mes_Cofins " & _
"FROM ((tbCalendario LEFT JOIN cstAcumCompras_credito ON (tbCalendario.ANO = cstAcumCompras_credito.ANO) AND (tbCalendario.MES = cstAcumCompras_credito.MES)) LEFT JOIN cstAcumTransportes ON (tbCalendario.ANO = cstAcumTransportes.ANO) AND (tbCalendario.MES = cstAcumTransportes.MES)) LEFT JOIN cstAcumVendas ON (tbCalendario.ANO = cstAcumVendas.ANO) AND (tbCalendario.MES = cstAcumVendas.MES) " & _
"GROUP BY tbCalendario.ANO, MESs " & _
"HAVING (tbCalendario.ANO='2017' AND MESs>=7) OR (tbCalendario.ANO>='2018') " & _
"ORDER BY tbCalendario.ANO, MESs ;")
Conn.Execute strSQL


lin = 1
credtransp = 0
Set rst = Db.OpenRecordset("tbResumo_Cofins")

Do Until rst.EOF
    If lin = 1 Then
    
    rst.Edit
        rst!CRED_MES_ANT = 0
        rst!CRED_TRANSPORTAR = rst!Saldo_Mes
        rst!SALDO = rst!Saldo_Mes
    rst.Update
    
    Else
        rst.MovePrevious
        credtransp = rst!CRED_TRANSPORTAR
        rst.MoveNext
        rst.Edit
        rst!CRED_MES_ANT = credtransp
            If rst!Saldo_Mes + credtransp <= 0 Then
            rst!CRED_TRANSPORTAR = 0
            Else
            rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp
            End If
            rst!SALDO = rst!Saldo_Mes + credtransp
        rst.Update
    End If
    
lin = lin + 1
rst.MoveNext
Loop
rst.Close


'ICMS_ST
'strSQL = ("INSERT INTO tbResumo_ICMS_ST ( ANO, MES, CRED, DEB, Saldo_Mes ) " & _
'"SELECT tbCalendario.ANO, CAST(tbCalendario.MES AS UNSIGNED) AS MESs, coalesce(Sum(cstAcumCompras.Valor_ICMS_ST),0)+coalesce(sum(cstAcumTransportes.Valor_ICMS_ST),0) AS CRED_ICMS_ST, coalesce(Sum(cstAcumVendas.Valor_ICMS_ST),0) AS DEB_ICMS_ST, coalesce(Sum(cstAcumCompras.Valor_ICMS_ST),0)+coalesce(SUM(cstAcumTransportes.Valor_ICMS_ST),0)-coalesce(Sum(cstAcumVendas.Valor_ICMS_ST),0) AS Saldo_Mes_ICMS_ST " & _
'"FROM ((tbCalendario LEFT JOIN cstAcumCompras ON (tbCalendario.ANO = cstAcumCompras.ANO) AND (tbCalendario.MES = cstAcumCompras.MES)) LEFT JOIN cstAcumTransportes ON (tbCalendario.ANO = cstAcumTransportes.ANO) AND (tbCalendario.MES = cstAcumTransportes.MES)) LEFT JOIN cstAcumVendas ON (tbCalendario.ANO = cstAcumVendas.ANO) AND (tbCalendario.MES = cstAcumVendas.MES) " & _
'"GROUP BY tbCalendario.ANO, MESs " & _
'"HAVING (tbCalendario.ANO='2017' AND MESs>=7) OR (tbCalendario.ANO>='2018');")
'Conn.Execute strSQL

'lin = 1
'credtransp = 0
'Set rst = Db.OpenRecordset("tbResumo_ICMS_ST")

'Do Until rst.EOF
'    If lin = 1 Then
'
'    rst.Edit
'        rst!CRED_MES_ANT = 0
'        rst!CRED_TRANSPORTAR = rst!Saldo_Mes
'        rst!SALDO = rst!Saldo_Mes
'    rst.Update
'
'    Else
'        rst.MovePrevious
'        credtransp = rst!CRED_TRANSPORTAR
'        rst.MoveNext
'        rst.Edit
'        rst!CRED_MES_ANT = credtransp
'            If rst!Saldo_Mes + credtransp <= 0 Then
'            rst!CRED_TRANSPORTAR = 0
'            Else
'            rst!CRED_TRANSPORTAR = rst!Saldo_Mes + credtransp
'            End If
'            rst!SALDO = rst!Saldo_Mes + credtransp
'        rst.Update
'    End If
    
'lin = lin + 1
'rst.MoveNext
'Loop
'rst.Close

'IRPJ E CSLL
Call Calc_IRPJ

Call DisconnectFromDataBase
'DoCmd.setwarnings (True)


End Sub

Sub Calcular_Custo_Medio()
Call ConnectToDataBase
'apenas para o cadastro inicial, não deve entrar na rotina de processamento
'DoCmd.setwarnings (False)

Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

Set rst = Db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.DescProd, tbCadProd.Unid, Sum(tbComprasDet.Qnt) AS Qnt, Sum(tbComprasDet.ValorTot) AS ValorTot " & _
"FROM tbCadProd INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd) " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.DescProd, tbCadProd.Unid;")

Do Until rst.EOF
strSQL = ("UPDATE tbCadProd SET tbCadProd.Estoque = " & Replace(rst!Qnt, ",", ".") & ", tbCadProd.CMed_Unit = " & Replace(rst!ValorTot / rst!Qnt, ",", ".") & " where tbCadProd.IDProd = " & rst!IDProd & ";")
Conn.Execute strSQL
rst.MoveNext
Loop

Set rst = Db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.DescProd, tbCadProd.Unid, Sum(tbVendasDet.Qnt) AS Qnt, Sum(tbVendasDet.ValorTot) AS ValorTot, tbVendas.TipoNF " & _
"FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.DescProd, tbCadProd.Unid, tbVendas.TipoNF " & _
"HAVING (((tbVendas.TipoNF)='0-ENTRADA'));")

Do Until rst.EOF
strSQL = ("UPDATE tbCadProd SET tbCadProd.Estoque = " & Replace(rst!Qnt, ",", ".") & ", tbCadProd.CMed_Unit = " & Replace(rst!ValorTot / rst!Qnt, ",", ".") & " where tbCadProd.IDProd = " & rst!IDProd & ";")
Conn.Execute strSQL
rst.MoveNext
Loop

Set rst = Db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.DescProd, tbCadProd.Unid, Sum(tbVendasDet.Qnt) AS Qnt, Sum(tbVendasDet.ValorTot) AS ValorTot, tbVendas.TipoNF " & _
"FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.DescProd, tbCadProd.Unid, tbVendas.TipoNF " & _
"HAVING (((tbVendas.TipoNF)='1-SAIDA'));")

Do Until rst.EOF
strSQL = ("UPDATE tbCadProd SET tbCadProd.Estoque = tbCadProd.Estoque - " & Replace(rst!Qnt, ",", ".") & " where tbCadProd.IDProd = " & rst!IDProd & ";")
Conn.Execute strSQL
rst.MoveNext
Loop


'calcular custo de itens de revenda com troca de código
'Cadastrar custo médio do que teve venda e não teve compra. Exemplo Barril x revenda de copos. Ou troca de códigos.
Set rst = Db.OpenRecordset("SELECT tbVendas.ANO, tbVendas.MES, tbVendas.DataEmissao, tbVendasDet.IDProd, tbVendasDet.Qnt, tbVendasDet.CustoMedio, tbVendasDet.VlrLiquido, tbCadProd.CMed_Unit, tbCadProd.DescProd, tbCadProd.LITROS, tbCadProd.IdProd_Revenda, tbcadprod_1.DescProd, tbcadprod_1.CMed_Unit as CMed_Unit_Original, tbcadprod_1.Estoque " & _
"FROM (tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) LEFT JOIN tbcadprod AS tbcadprod_1 ON tbCadProd.IdProd_Revenda = tbcadprod_1.IDProd " & _
"WHERE (((tbVendasDet.CustoMedio)=0) AND ((tbVendasDet.CFOP)<>'1604'));")

Do Until rst.EOF
strSQL = ("UPDATE tbCadProd SET tbCadProd.Estoque = 0, tbCadProd.CMed_Unit = " & Replace(rst!CMed_Unit_Original * rst!Litros, ",", ".") & " where tbCadProd.IDProd = " & rst!IDProd & ";")
Conn.Execute strSQL
rst.MoveNext
Loop



Call DisconnectFromDataBase

'DoCmd.setwarnings (True)
End Sub


Private Sub Calc_IRPJ()

Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

Dim lin As Integer
Dim credtransp As Double

Dim cICMS_Ciap As Double


Call ConnectToDataBase
'IRPJ E CSLL
strSQL = ("delete from tbResumo_IRPJ_CSLL;")
Conn.Execute strSQL

strSQL = ("delete from tbResumoVendas;")
Conn.Execute strSQL

strSQL = ("INSERT INTO tbResumoVendas (ANO, MES, ValorTot, Valor_ICMS, Valor_PIS, Valor_Cofins, Valor_IPI, Valor_ICMS_ST) select * FROM cstAcumVendas;")
Conn.Execute strSQL

strSQL = ("INSERT INTO tbResumo_IRPJ_CSLL ( ANOMES, ANO, MES, APURACAO, REGIME, ALIQ_BC_IRPJ, ALIQ_IRPJ, ALIQ_BC_CSLL, ALIQ_CSLL ) " & _
"SELECT CONCAT(tbAliq_IRPJ_CSLL.ANO,LPAD(tbAliq_IRPJ_CSLL.MES,2,'0')), tbAliq_IRPJ_CSLL.ANO, tbAliq_IRPJ_CSLL.MES, tbAliq_IRPJ_CSLL.APURACAO, tbAliq_IRPJ_CSLL.REGIME, tbAliq_IRPJ_CSLL.ALIQ_BC_IRPJ, tbAliq_IRPJ_CSLL.ALIQ_IRPJ, tbAliq_IRPJ_CSLL.ALIQ_BC_CSLL, tbAliq_IRPJ_CSLL.ALIQ_CSLL " & _
"FROM tbAliq_IRPJ_CSLL INNER JOIN tbCalendario ON (tbAliq_IRPJ_CSLL.MES = tbCalendario.MES) AND (tbAliq_IRPJ_CSLL.ANO = tbCalendario.ANO);")
Conn.Execute strSQL

strSQL = ("UPDATE tbResumo_IRPJ_CSLL set Fat_Bruto = 0;")
Conn.Execute strSQL


Dim cAno As Integer
Dim cMes As Integer
Dim cFatBruto As Double

Set rst = Db.OpenRecordset("tbResumo_IRPJ_CSLL")

Do Until rst.EOF
    Select Case rst!APURACAO
    Case Is = "MENSAL"
       cAno = rst!ANO
       cMes = rst!MES
       'SELECT Sum(cstAcumVendas.ValorTot-valor_icms-valor_pis-valor_cofins-valor_IPI-valor_ICMS_ST) AS FatBruto FROM cstAcumVendas WHERE (((cstAcumVendas.ANO)="2018"));

       Set rstFat = Db.OpenRecordset("SELECT Sum(ValorTot-valor_icms-valor_pis-valor_cofins-valor_IPI-valor_ICMS_ST) AS FatBruto FROM cstAcumVendas WHERE ANO='" & cAno & "' AND MES=" & cMes & ";")
       Case Is = "TRIMESTRAL"
       Select Case rst!MES
       'Q1'
       Case Is = 3
       cAno = rst!ANO
       strSQL = "UPDATE tbResumo_IRPJ_CSLL as Q1 INNER JOIN (select ANOMES, sum(FatBruto) as FatBruto from (SELECT CONCAT(ANO,'03') AS ANOMES, SUM(FatBruto) as FatBruto FROM (SELECT concat(ANO,LPAD(MES,2,'0')) AS ANOMES, ANO, Sum(ValorTot) as FatBruto FROM cstAcumVendas WHERE ANO= " & cAno & " AND (MES=1 Or MES=2 Or MES=3) GROUP BY ANO, ANOMES) as Q2 group by ANOMES) AS Q2 group by ANOMES) as Q2 ON Q1.ANOMES = Q2.ANOMES SET Q1.FAT_BRUTO = Q2.FatBruto;"
       Conn.Execute strSQL
       'Q2
       Case Is = 6
       cAno = rst!ANO
       strSQL = "UPDATE tbResumo_IRPJ_CSLL as Q1 INNER JOIN (select ANOMES, sum(FatBruto) as FatBruto from (SELECT CONCAT(ANO,'06') AS ANOMES, SUM(FatBruto) as FatBruto FROM (SELECT concat(ANO,LPAD(MES,2,'0')) AS ANOMES, ANO, Sum(ValorTot) as FatBruto FROM cstAcumVendas WHERE ANO=" & cAno & " AND (MES=4 Or MES=5 Or MES=6) GROUP BY ANO, ANOMES) as Q2 group by ANOMES) AS Q2 group by ANOMES) as Q2 ON Q1.ANOMES = Q2.ANOMES SET Q1.FAT_BRUTO = Q2.FatBruto;"
       Conn.Execute strSQL
       'Q3
       Case Is = 9
       cAno = rst!ANO
       strSQL = "UPDATE tbResumo_IRPJ_CSLL as Q1 INNER JOIN (select ANOMES, sum(FatBruto) as FatBruto from (SELECT CONCAT(ANO,'09') AS ANOMES, SUM(FatBruto) as FatBruto FROM (SELECT concat(ANO,LPAD(MES,2,'0')) AS ANOMES, ANO, Sum(ValorTot) as FatBruto FROM cstAcumVendas WHERE ANO=" & cAno & " AND (MES=7 Or MES=8 Or MES=9) GROUP BY ANO, ANOMES) as Q2 group by ANOMES) AS Q2 group by ANOMES) as Q2 ON Q1.ANOMES = Q2.ANOMES SET Q1.FAT_BRUTO = Q2.FatBruto;"
       Conn.Execute strSQL
       'Q4
       Case Is = 12
       cAno = rst!ANO
       strSQL = "UPDATE tbResumo_IRPJ_CSLL as Q1 INNER JOIN (select ANOMES, sum(FatBruto) as FatBruto from (SELECT CONCAT(ANO,'12') AS ANOMES, SUM(FatBruto) as FatBruto FROM (SELECT concat(ANO,LPAD(MES,2,'0')) AS ANOMES, ANO, Sum(ValorTot) as FatBruto FROM cstAcumVendas WHERE ANO=" & cAno & " AND (MES=10 Or MES=11 Or MES=12) GROUP BY ANO, ANOMES) as Q2 group by ANOMES) AS Q2 group by ANOMES) as Q2 ON Q1.ANOMES = Q2.ANOMES SET Q1.FAT_BRUTO = Q2.FatBruto;"
       Conn.Execute strSQL
       End Select
    Case Is = "ANUAL"
       'cAno = rst!ANO
       'Set rstFat = Db.OpenRecordset("SELECT Sum(tbResumoVendas.ValorTot-valor_icms-valor_pis-valor_cofins-valor_IPI-valor_ICMS_ST) AS FatBruto FROM tbResumoVendas HAVING (((tbResumoVendas.ANO)='" & cAno & "'));")
    End Select
    
    
rst.MoveNext
Loop

'CORRIGE BASE CALC = 0 - DÁ PAU NO EFD CONTRIBUIÇÕES
strSQL = ("Update tbCompras set ICMS_BaseCalc = VlrTOTALNF where ICMS_BaseCalc = 0;")
Conn.Execute strSQL
strSQL = ("Update tbComprasDet set BaseCalculo = ValorTot where BaseCalculo = 0;")
Conn.Execute strSQL

'BASE CALC
strSQL = ("UPDATE tbResumo_IRPJ_CSLL set BC_IRPJ = FAT_BRUTO * ALIQ_BC_IRPJ where REGIME = 'PRESUMIDO'")
Conn.Execute strSQL
strSQL = ("UPDATE tbResumo_IRPJ_CSLL set BC_CSLL = FAT_BRUTO * ALIQ_BC_CSLL where REGIME = 'PRESUMIDO'")
Conn.Execute strSQL
'VALOR IRPJ E CSLL
strSQL = ("UPDATE tbResumo_IRPJ_CSLL set VL_IRPJ = BC_IRPJ * ALIQ_IRPJ where REGIME = 'PRESUMIDO'")
Conn.Execute strSQL
strSQL = ("UPDATE tbResumo_IRPJ_CSLL set VL_CSLL = BC_CSLL * ALIQ_CSLL where REGIME = 'PRESUMIDO'")
Conn.Execute strSQL
'IRPJ E CSLL


End Sub






