Attribute VB_Name = "modGeraArquivo_GIA"
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
Conn.Open "DRIVER={MySQL ODBC 8.0 ANSI Driver}" _
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
Public Function Gerar_GIA(cDtIni As String, cDtFim As String, clocal As String, cDtINI_Contabil As String)

ConnectToDataBase

'ARQUIVO GIA
'DoCmd.setwarnings (False)

'EXPORTAR ARQUIVO TXT
Dim iArq As Long
iArq = FreeFile

Open clocal & "\GIA_" & month(cDtIni) & "_" & year(cDtIni) & ".prf" For Output As iArq

'Print #iArq, c0000 & Chr(13); c0001 & Chr(13); c0005 & Chr(13); c0100 & Chr(13); c0150 & Chr(13) & c0190 & Chr(13) & c0200 & Chr(13) & c0300 & Chr(13) & c0305 & Chr(13) & c0400 & Chr(13) & c0500 & Chr(13); c0600 & Chr(13) & c0990 & Chr(13) & cC001 & Chr(13) & cC100 & Chr(13) & cC170 & Chr(13) & cC190 & Chr(13) & cC500 & Chr(13) & cC501 & Chr(13) & cC990 & Chr(13) & cD001 & Chr(13) & cD190 & Chr(13) & cD990 & Chr(13) & cE001 & Chr(13) & cE100
'Print #iArq, c0000
Dim cSTR_DtINI As String
Dim cSTR_DtFIM As String

cSTR_DtINI = Replace(Format(cDtIni, "dd/mm/yyyy"), "/", "")
cSTR_DtFIM = Replace(Format(cDtFim, "dd/mm/yyyy"), "/", "")


cDtIni = Format(cDtIni, "mm/dd/yyyy")
cDtFim = Format(cDtFim, "mm/dd/yyyy")

Dim cAno As String
Dim cMes As String

cAno = Right(Left(cSTR_DtINI, 8), 4)
cMes = Int(Right(Left(cSTR_DtINI, 4), 2))

If cMes < 10 Then
cMes = "0" & cMes
Else
End If



Dim Db As Database
Set Db = CurrentDb()

Dim rsEmpresa As DAO.Recordset
Set rsEmpresa = Db.OpenRecordset("tbEmpresa")

Dim rsContador As DAO.Recordset
Set rsContador = Db.OpenRecordset("tbContador")


Dim lCount As Integer
Dim lCR01 As Integer
Dim lCR05 As Integer
Dim lCR07 As Integer
Dim lCR010 As Integer
Dim lCR020 As Integer
Dim lCR030 As Integer
Dim lCR031 As Integer


lCount = 0
lCR01 = 0
lCR05 = 0
lCR07 = 0
lCR010 = 0
lCR020 = 0
lCR030 = 0
lCR031 = 0

'CR=01 - REGISTRO MESTRE
Dim r1_1 As String * 2
Dim r1_2 As String * 2
Dim r1_3 As String * 8
Dim r1_4 As String * 6
Dim r1_5 As String * 4
Dim r1_6 As String * 4
Dim r1_7 As String * 4



r1_1 = "01"
r1_2 = "01"
r1_3 = year(Now) & Format(month(Now), "00") & Format(Day(Now), "00")
r1_4 = Format(hour(Now), "00") & Format(minute(Now), "00") & Format(second(Now), "00")
r1_5 = "0000"
r1_6 = "0210" 'versao
r1_7 = "" 'filhos CR=05

lCR01 = lCR01 + 1


'cCR01 = r1_1 & r1_2 & r1_3 & r1_4 & r1_5 & r1_6 & r1_7 '& Chr(13) + Chr(10)
'Print #iArq, cCR01
'cCount = cCount + 1


'CR-05 - CABEÇALHO DO DOCUMENTO FISCAL
Dim r2_1 As String * 2
Dim r2_2 As String * 12
Dim r2_3 As String * 14
Dim r2_4 As String * 7
Dim r2_5 As String * 2
Dim r2_6 As String * 6
Dim r2_7 As String * 6
Dim r2_8 As String * 2
Dim r2_9 As String * 1
Dim r2_10 As String * 1
Dim r2_11 As String * 15
Dim r2_12 As String * 15
Dim r2_13 As String * 14
Dim r2_14 As String * 1
Dim r2_15 As String * 15
Dim r2_16 As String * 32
Dim r2_17 As String * 4
Dim r2_18 As String * 4
Dim r2_19 As String * 4
Dim r2_20 As String * 4
Dim r2_21 As String * 4
Dim r2_22 As String * 2
Dim r2_23 As String * 1

Dim rsICMS_Saldo As DAO.Recordset
Set rsICMS_Saldo = Db.OpenRecordset("select * from tbResumo_ICMS where ano = '" & cAno & "' and mes = " & Val(cMes) & "")

Dim rsICMS_ST_Saldo As DAO.Recordset
Set rsICMS_ST_Saldo = Db.OpenRecordset("select * from tbResumo_ICMS_ST where ano = '" & cAno & "' and mes = " & Val(cMes) & "")



r2_1 = "05"
r2_2 = rsEmpresa!IE
r2_3 = rsEmpresa!CNPJ
'r2_4 = rsEmpresa!CNAE ' preencher com zeros campo inutil
r2_4 = "0000000"
r2_5 = "01" 'RPA - Regime Periodico de Apuracao
r2_6 = cAno & Format(cMes, "00") 'ANO E MES DA GIA
r2_7 = "000000" ' Ref Inicial preencher com zeros se RPA
r2_8 = "01" 'Tipo da GIA 01 Normal
r2_9 = "0" 'se há movimento = 0 senão 1
r2_10 = "0" 'indica se já foi transmitido
r2_11 = Format(Replace(Format(Round(rsICMS_Saldo!CRED_MES_ANT, 2), "0.00"), ",", ""), "000000000000000") 'saldo credor periodo anterior Alinhar à direita e preencher com ZEROS à esquerda
r2_12 = Format(Replace(Format(Round(rsICMS_ST_Saldo!CRED_MES_ANT, 2), "0.00"), ",", ""), "000000000000000") 'saldo credor periodo anterior ST
r2_13 = rsEmpresa!CNPJ
r2_14 = "0"
r2_15 = "000000000000000" ' ICMS Fixado para o período
r2_16 = "00000000000000000000000000000000"
r2_17 = "" 'qt filhos CR7
r2_18 = "" 'qt filhos CR10
r2_19 = "" 'qt filhos CR20
r2_20 = "" 'qt filhos CR30
r2_21 = "" 'qt filhos CR31

lCR05 = lCR05 + 1

'cCR02 = r2_1 & r2_2 & r2_3 & r2_4 & r2_5 & r2_6 & r2_7 & r2_8 & r2_9 & r2_10 & r2_11 & r2_12 & r2_13 & r2_14 & r2_15 & r2_16 & r2_17 & r2_18 & r2_19 & r2_20 & r2_21 '& Chr(13) + Chr(10)
'Print #iArq, cCR02
'cCount = cCount + 1


'CR-07 - DETALHES PAGAMENTOS
'nao obrigatório, referente a data de vencimento
lCR07 = 0

'CR-10 - Detalhes CFOPs
Dim r10_1 As String * 2
Dim r10_2 As String * 6
Dim r10_3 As String * 15
Dim r10_4 As String * 15
Dim r10_5 As String * 15
Dim r10_6 As String * 15
Dim r10_7 As String * 15
Dim r10_8 As String * 15
Dim r10_9 As String * 15
Dim r10_10 As String * 15
Dim r10_11 As String * 15
Dim r10_12 As String * 4


strSQL = ("delete from tbResumoGIA")
Conn.Execute strSQL
strSQL = ("INSERT INTO tbResumoGIA " & _
"select * from( " & _
"SELECT tbCompras.ANO, tbCompras.MES, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, Sum(tbComprasDet.Valor_ICMS) AS ValorICMS, Sum(tbComprasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbComprasDet.ValorTot) AS ValorTot " & _
"FROM tbCadProd INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd) " & _
"WHERE tbComprasDet.LancFiscal = 'CREDITO' " & _
"GROUP BY tbCompras.ANO, tbCompras.MES, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
"HAVING tbCompras.ANO = '" & cAno & "' And tbCompras.MES = " & Int(cMes) & " " & _
"union " & _
"SELECT tbTransportes.ANO, tbTransportes.MES, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC, Sum(tbTransportes.ValorICMS) AS ValorICMS, Sum(tbTransportes.BaseCalcICMS) AS BaseCalcICMS, Sum(tbTransportes.ValorTotalServico) AS ValorTot " & _
"FROM tbTransportes INNER JOIN tbTransportesDet ON tbTransportes.ID = tbTransportesDet.ID_Transporte " & _
"WHERE tbTransportes.Creditavel = 'SIM' And tbTransportes.LancFiscal = 'CREDITO' " & _
"GROUP BY tbTransportes.ANO, tbTransportes.MES, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC " & _
"HAVING tbTransportes.ANO = '" & cAno & "' And tbTransportes.MES = " & Int(cMes) & " " & _
"union " & _
"SELECT tbVendas.ANO, tbVendas.MES, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, Sum(tbVendasDet.Valor_ICMS) AS ValorICMS, Sum(tbVendasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbVendasDet.ValorTot) AS ValorTot " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE tbVendasDet.LancFiscal = 'DEBITO' AND `status` = 'ATIVO' " & _
"GROUP BY tbVendas.ANO, tbVendas.MES, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
"HAVING tbVendas.ANO = '" & cAno & "' And tbVendas.MES = " & Int(cMes) & " " & _
"union " & _
"SELECT tbVendas.ANO, tbVendas.MES, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, Sum(tbVendasDet.Valor_ICMS) AS ValorICMS, Sum(tbVendasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbVendasDet.ValorTot) AS ValorTot " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE CFOP_ESCRITURADA = '1604' AND `status` = 'ATIVO' " & _
"GROUP BY tbVendas.ANO, tbVendas.MES, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
"HAVING tbVendas.ANO = '" & cAno & "' And tbVendas.MES = " & Int(cMes) & ") AS Q1 ORDER BY CFOP_ESCRITURADA ")
Conn.Execute strSQL

' Esse imobilizado é referente ao credito devido, alterei com o credito de NF de entrada para bater com a EFD
'"SELECT tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.CFOP AS CFOP_ESCRITURADA, tbImobilizado.`CFOP Desc` AS CFOP_ESC_DESC, Sum(tbImobilizado.Valor_ICMS) AS Valor_ICMS, Sum(tbImobilizado.BaseCalculo) AS BaseCalcICMS, Sum(tbImobilizado.ValorTot) AS ValorTot " & _
'"FROM tbImobilizado " & _
'"GROUP BY tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.CFOP, tbImobilizado.`CFOP Desc` " & _
'"HAVING tbImobilizado.ANO = '" & cAno & "' And tbImobilizado.MES = " & Int(cMes) & " " & _



'PARECE QUE O ODBC NÃO AGUENTA UMA STRING TÃO GRANDE
'Dim rsCFOP_ESC As dao.Recordset
'Set rsCFOP_ESC = db.OpenRecordset("" & _
'"SELECT * FROM ( " & _
'"SELECT tbCompras.ANO, tbCompras.MES, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, Sum(tbComprasDet.Valor_ICMS) AS ValorICMS, Sum(tbComprasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbComprasDet.ValorTot) AS ValorTot " & _
'"FROM tbCadProd INNER JOIN (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd) " & _
'"WHERE tbComprasDet.LancFiscal = 'CREDITO' " & _
'"GROUP BY tbCompras.ANO, tbCompras.MES, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
'"HAVING tbCompras.ANO = '" & cAno & "' And tbCompras.MES = " & Int(cMes) & " " & _
'"UNION " & _
'"SELECT tbTransportes.ANO, tbTransportes.MES, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC, Sum(tbTransportes.ValorICMS) AS ValorICMS, Sum(tbTransportes.BaseCalcICMS) AS BaseCalcICMS, Sum(tbTransportes.ValorTotalServico) AS ValorTot " & _
'"FROM tbTransportes INNER JOIN tbTransportesDet ON tbTransportes.ID = tbTransportesDet.ID_Transporte " & _
'"WHERE tbTransportes.Creditavel = 'SIM' And tbTransportes.LancFiscal = 'CREDITO' " & _
'"GROUP BY tbTransportes.ANO, tbTransportes.MES, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC " & _
'"HAVING tbTransportes.ANO = '" & cAno & "' And tbTransportes.MES = " & Int(cMes) & " " & _
'"SELECT * FROM ( " & _
'"SELECT tbVendas.ANO, tbVendas.MES, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, Sum(tbVendasDet.Valor_ICMS) AS ValorICMS, Sum(tbVendasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbVendasDet.ValorTot) AS ValorTot " & _
'"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
'"WHERE tbVendasDet.LancFiscal = 'DEBITO' AND [STATUS] = 'ATIVO' " & _
'"GROUP BY tbVendas.ANO, tbVendas.MES, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
'"HAVING tbVendas.ANO = '" & cAno & "' And tbVendas.MES = " & Int(cMes) & " " & _
'"UNION " & _
'"SELECT tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.CFOP AS CFOP_ESCRITURADA, tbImobilizado.[CFOP Desc] AS CFOP_ESC_DESC, Sum(tbImobilizado.Valor_ICMS) AS Valor_ICMS, Sum(tbImobilizado.BaseCalculo) AS BaseCalcICMS, Sum(tbImobilizado.ValorTot) AS ValorTot " & _
'"FROM tbImobilizado " & _
'"GROUP BY tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.CFOP, tbImobilizado.[CFOP Desc] " & _
'"HAVING tbImobilizado.ANO = '" & cAno & "' And tbImobilizado.MES = " & Int(cMes) & " " & _
'") AS Q1 ORDER BY CFOP_ESCRITURADA")

'se o ICMS = 0 dá erro no layout do arquivo
'mas se remove o registro da inconsistência na GIA EFD PUTA QUE O PARIU
strSQL = ("delete from tbResumoGIA where ValorICMS = 0")
Conn.Execute strSQL


Dim rsCFOP_ESC As DAO.Recordset
Set rsCFOP_ESC = Db.OpenRecordset("select * from tbResumoGIA order by int(CFOP_ESCRITURADA)")


zCR014 = "NAO"
Do Until rsCFOP_ESC.EOF = True
If Left(rsCFOP_ESC!CFOP_ESCRITURADA, 1) = 2 Or Left(rsCFOP_ESC!CFOP_ESCRITURADA, 1) = 6 Then
zCR014 = "SIM"
Else
End If
lCR010 = lCR010 + 1
rsCFOP_ESC.MoveNext
Loop


If zCR014 = "SIM" Then
Else
GoTo CR014_NAO
End If

'CR=14– Detalhes Interestaduais - Quando houver vendas para outras UFs
Dim r14_1 As String * 2
Dim r14_2 As String * 2
Dim r14_3 As String * 15
Dim r14_4 As String * 15
Dim r14_5 As String * 15
Dim r14_6 As String * 15
Dim r14_7 As String * 15
Dim r14_8 As String * 15
Dim r14_9 As String * 15
Dim r14_10 As String * 15
Dim r14_11 As String * 15
Dim r14_12 As String * 1
Dim r14_13 As String * 4
Dim r14_14 As String * 2

r14_1 = ""
r14_2 = ""
r14_3 = ""
r14_4 = ""
r14_5 = ""
r14_6 = ""
r14_7 = ""
r14_8 = ""
r14_9 = ""
r14_10 = ""
r14_11 = ""
r14_12 = ""
r14_13 = ""
r14_14 = ""

CR014_NAO:


'CR=18–ZFM/ALC - Quando houver transações para zona franca
lCR018 = 0
'CR=20–Ocorrências - Não preencher
lCR020 = 0
'CR=25–IEs - Associadas as ocorrencias
'CR=30-DIPAM-B
lCR030 = 0
'CR=31–Registro de Exportação
lCR031 = 0



'GERAR LINHAS

'ATUALIZAR CONTADORES
r1_7 = Format(lCR01, "0000")
r2_17 = Format(lCR07, "0000")
r2_21 = Format(lCR031, "0000")
r2_20 = Format(lCR030, "0000")
r2_19 = Format(lCR020, "0000")
r2_18 = Format(lCR010, "0000")

'ATUALIZAR CONTADORES

cCR01 = r1_1 & r1_2 & r1_3 & r1_4 & r1_5 & r1_6 & r1_7
Print #iArq, cCR01
cCount = cCount + 1

cCR02 = r2_1 & r2_2 & r2_3 & r2_4 & r2_5 & r2_6 & r2_7 & r2_8 & r2_9 & r2_10 & r2_11 & r2_12 & r2_13 & r2_14 & r2_15 & r2_16 & r2_17 & r2_18 & r2_19 & r2_20 & r2_21
Print #iArq, cCR02
cCount = cCount + 1

Dim rsCFOP_014 As DAO.Recordset

If lCR010 > 0 Then
rsCFOP_ESC.MoveFirst
Else
End If
Do Until rsCFOP_ESC.EOF = True
r10_1 = "10" 'Código do registro
r10_2 = Format(rsCFOP_ESC!CFOP_ESCRITURADA, "0000") & "00" 'CFOP da transação
r10_3 = Format(Replace(Format(Round(rsCFOP_ESC!ValorTot, 2), "0.00"), ",", ""), "000000000000000")  'Valor Contabil da transação
r10_4 = Format(Replace(Format(Round(rsCFOP_ESC!BaseCalcICMS, 2), "0.00"), ",", ""), "000000000000000") 'Base de Calculo
r10_5 = Format(Replace(Format(Round(rsCFOP_ESC!ValorICMS, 2), "0.00"), ",", ""), "000000000000000") 'Imposto Creditado ou Debitado
r10_6 = "000000000000000" 'Isentas e Não Tributadas
r10_7 = "000000000000000"  'Outros valores
r10_8 = "000000000000000"   'Imposto Retido por Substituição Tributária
r10_9 = "000000000000000"  'Imposto lançado para contribuinte Do tipo Substituto, responsável pelo recolhimento do imposto
r10_10 = "000000000000000" 'Imposto Retido por Substituição Tributária (Substituído)
r10_11 = "000000000000000" 'Outros Impostos
r10_12 = "0000" 'Registros Q14
    
    If Left(rsCFOP_ESC!CFOP_ESCRITURADA, 1) = 2 Or Left(rsCFOP_ESC!CFOP_ESCRITURADA, 1) = 6 Then
        strSQL = ("delete from tbResumoGIA_Det")
        Conn.Execute strSQL
        strSQL = ("insert into tbResumoGIA_Det SELECT * FROM " & _
        "(SELECT tbCompras.ANO, tbCompras.MES, tbFornecedor.UF, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, Sum(tbComprasDet.Valor_ICMS) AS ValorICMS, Sum(tbComprasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbComprasDet.ValorTot) AS ValorTot " & _
        "FROM tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
        "WHERE tbComprasDet.LancFiscal = 'CREDITO' " & _
        "GROUP BY tbCompras.ANO, tbCompras.MES, tbFornecedor.UF, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
        "HAVING tbCompras.ANO = '" & cAno & "' And tbCompras.MES = " & Int(cMes) & " UNION " & _
        "SELECT tbTransportes.ANO, tbTransportes.MES, tbFornecedor.UF, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC, Sum(tbTransportes.ValorICMS) AS ValorICMS, Sum(tbTransportes.BaseCalcICMS) AS BaseCalcICMS, Sum(tbTransportes.ValorTotalServico) AS ValorTot " & _
        "FROM tbFornecedor INNER JOIN (tbTransportes INNER JOIN tbTransportesDet ON tbTransportes.ID = tbTransportesDet.ID_Transporte) ON tbFornecedor.IDFor = tbTransportes.ID_Emit " & _
        "WHERE tbTransportes.Creditavel = 'SIM' And tbTransportes.LancFiscal = 'CREDITO' " & _
        "GROUP BY tbTransportes.ANO, tbTransportes.MES, tbFornecedor.UF, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC " & _
        "HAVING tbTransportes.ANO = '" & cAno & "' And tbTransportes.MES = " & Int(cMes) & " UNION " & _
        "SELECT tbVendas.ANO, tbVendas.MES, tbCliente.UF, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, Sum(tbVendasDet.Valor_ICMS) AS ValorICMS, Sum(tbVendasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbVendasDet.ValorTot) AS ValorTot " & _
        "FROM tbCliente INNER JOIN (tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda) ON tbCliente.IDCliente = tbVendas.IdCliente " & _
        "WHERE tbVendasDet.LancFiscal = 'DEBITO' AND `STATUS` = 'ATIVO' " & _
        "GROUP BY tbVendas.ANO, tbVendas.MES, tbCliente.UF, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
        "HAVING tbVendas.ANO = '" & cAno & "' And tbVendas.MES = " & Int(cMes) & " UNION " & _
        "SELECT tbImobilizado.ANO, tbImobilizado.MES, tbFornecedor.UF, tbImobilizado.CFOP AS CFOP_ESCRITURADA, tbImobilizado.`CFOP Desc` AS CFOP_ESC_DESC, Sum(tbImobilizado.Valor_ICMS) AS Valor_ICMS, Sum(tbImobilizado.BaseCalculo) AS BaseCalcICMS, Sum(tbImobilizado.ValorTot) AS ValorTot " & _
        "FROM tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor INNER JOIN tbImobilizado ON tbCompras.ChaveNF = tbImobilizado.ChaveNFe " & _
        "GROUP BY tbImobilizado.ANO, tbImobilizado.MES, tbFornecedor.UF, tbImobilizado.CFOP, tbImobilizado.`CFOP Desc` " & _
        "HAVING tbImobilizado.ANO = '" & cAno & "' And tbImobilizado.MES = " & Int(cMes) & " " & _
        ") AS Q1 WHERE CFOP_ESCRITURADA = '" & rsCFOP_ESC!CFOP_ESCRITURADA & "' ORDER BY UF;")
        Conn.Execute strSQL
        
        Set rsCFOP_014 = Db.OpenRecordset("select * from tbResumoGIA_Det order by UF")
       
        'Set rsCFOP_014 = db.OpenRecordset("" & _
        '"SELECT * FROM (SELECT tbCompras.ANO, tbCompras.MES, tbFornecedor.UF, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, Sum(tbComprasDet.Valor_ICMS) AS ValorICMS, Sum(tbComprasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbComprasDet.ValorTot) AS ValorTot " & _
        '"FROM tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
        '"WHERE (((tbComprasDet.LancFiscal) = 'CREDITO')) " & _
        '"GROUP BY tbCompras.ANO, tbCompras.MES, tbFornecedor.UF, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
        '"HAVING (((tbCompras.ANO) = '" & cAno & "') And ((tbCompras.MES) = " & Int(cMes) & ")) " & _
        '"UNION " & _
        '"SELECT tbTransportes.ANO, tbTransportes.MES, tbFornecedor.UF, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC, Sum(tbTransportes.ValorICMS) AS ValorICMS, Sum(tbTransportes.BaseCalcICMS) AS BaseCalcICMS, Sum(tbTransportes.ValorTotalServico) AS ValorTot " & _
        '"FROM tbFornecedor INNER JOIN (tbTransportes INNER JOIN tbTransportesDet ON tbTransportes.ID = tbTransportesDet.ID_Transporte) ON tbFornecedor.IDFor = tbTransportes.ID_Emit " & _
        '"WHERE (((tbTransportes.Creditavel) = 'SIM') And ((tbTransportes.LancFiscal) = 'CREDITO')) " & _
        '"GROUP BY tbTransportes.ANO, tbTransportes.MES, tbFornecedor.UF, tbTransportes.CFOP_ESCRITURADA, tbTransportes.CFOP_ESC_DESC " & _
        '"HAVING (((tbTransportes.ANO) = '" & cAno & "') And ((tbTransportes.MES) = " & Int(cMes) & ")) " & _
        '"UNION " & _
        '"SELECT tbVendas.ANO, tbVendas.MES, tbCliente.UF, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, Sum(tbVendasDet.Valor_ICMS) AS ValorICMS, Sum(tbVendasDet.BaseCalculo) AS BaseCalcICMS, Sum(tbVendasDet.ValorTot) AS ValorTot " & _
        '"FROM tbCliente INNER JOIN (tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda) ON tbCliente.IDCliente = tbVendas.IdCliente " & _
        '"WHERE (((tbVendasDet.LancFiscal) = 'DEBITO' AND [STATUS] = 'ATIVO')) " & _
        '"GROUP BY tbVendas.ANO, tbVendas.MES, tbCliente.UF, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
        '"HAVING (((tbVendas.ANO) = '" & cAno & "') And ((tbVendas.MES) = " & Int(cMes) & ")) " & _
        '"UNION " & _
        '"'SELECT tbImobilizado.ANO, tbImobilizado.MES, tbFornecedor.UF, tbImobilizado.CFOP AS CFOP_ESCRITURADA, tbImobilizado.[CFOP Desc] AS CFOP_ESC_DESC, Sum(tbImobilizado.Valor_ICMS) AS Valor_ICMS, Sum(tbImobilizado.BaseCalculo) AS BaseCalcICMS, Sum(tbImobilizado.ValorTot) AS ValorTot " & _
        '"FROM (tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor) INNER JOIN tbImobilizado ON tbCompras.ChaveNF = tbImobilizado.ChaveNFe " & _
        '"GROUP BY tbImobilizado.ANO, tbImobilizado.MES, tbFornecedor.UF, tbImobilizado.CFOP, tbImobilizado.[CFOP Desc] " & _
        '"HAVING (((tbImobilizado.ANO) = '" & cAno & "') And ((tbImobilizado.MES) = " & Int(cMes) & "))) AS Q1 WHERE CFOP_ESCRITURADA = '" & rsCFOP_ESC!CFOP_ESCRITURADA & "' ORDER BY UF;")

        cLinRsCFOP_014 = 0
        Do Until rsCFOP_014.EOF = True
        cLinRsCFOP_014 = cLinRsCFOP_014 + 1
        rsCFOP_014.MoveNext
        Loop
        
        r10_12 = Format(cLinRsCFOP_014, "0000") 'Quantidade de registros CR=14
         
        cCR10 = r10_1 & r10_2 & r10_3 & r10_4 & r10_5 & r10_6 & r10_7 & r10_8 & r10_9 & r10_10 & r10_11 & r10_12
        Print #iArq, cCR10
        
        rsCFOP_014.MoveFirst
        
        Do Until rsCFOP_014.EOF = True
        r14_1 = "14"
        Select Case rsCFOP_014!UF
        Case Is = "AC"
        r14_2 = "01" 'UF
        Case Is = "AL"
        r14_2 = "02" 'UF
        Case Is = "AP"
        r14_2 = "03" 'UF
        Case Is = "AM"
        r14_2 = "04" 'UF
        Case Is = "BA"
        r14_2 = "05" 'UF
        Case Is = "CE"
        r14_2 = "06" 'UF
        Case Is = "DF"
        r14_2 = "07" 'UF
        Case Is = "ES"
        r14_2 = "08" 'UF
        Case Is = "GO"
        r14_2 = "10" 'UF
        Case Is = "MA"
        r14_2 = "12" 'UF
        Case Is = "MT"
        r14_2 = "13" 'UF
        Case Is = "MS"
        r14_2 = "28" 'UF
        Case Is = "MG"
        r14_2 = "14" 'UF
        Case Is = "PA"
        r14_2 = "15" 'UF
        Case Is = "PB"
        r14_2 = "16" 'UF
        Case Is = "PR"
        r14_2 = "17" 'UF
        Case Is = "PE"
        r14_2 = "18" 'UF
        Case Is = "PI"
        r14_2 = "19" 'UF
        Case Is = "RJ"
        r14_2 = "22" 'UF
        Case Is = "RN"
        r14_2 = "20" 'UF
        Case Is = "RS"
        r14_2 = "21" 'UF
        Case Is = "RO"
        r14_2 = "23" 'UF
        Case Is = "RR"
        r14_2 = "24" 'UF
        Case Is = "SC"
        r14_2 = "25" 'UF
        Case Is = "SP"
        r14_2 = "26" 'UF
        Case Is = "SE"
        r14_2 = "27" 'UF
        Case Is = "TO"
        r14_2 = "29" 'UF
        
        End Select
        r14_3 = Format(Replace(Format(Round(rsCFOP_014!ValorTot, 2), "0.00"), ",", ""), "000000000000000") 'Valor Contabil Contribuinte
        r14_4 = Format(Replace(Format(Round(rsCFOP_014!BaseCalcICMS, 2), "0.00"), ",", ""), "000000000000000") 'Base de Calculo
        r14_5 = "000000000000000" 'Valor contabil Nao contribuinte - PF
        r14_6 = "000000000000000" 'Base Calculo nao contribuinte - PF
        r14_7 = Format(Replace(Format(Round(rsCFOP_014!ValorICMS, 2), "0.00"), ",", ""), "000000000000000") 'Imposto
        r14_8 = "000000000000000" 'outros valores
        r14_9 = "000000000000000" 'ICMS ST
        r14_10 = "000000000000000" 'Petrole e energia
        r14_11 = "000000000000000" 'outros ICMS ST
        r14_12 = "0" 'Indica se há zona franca - Não venda pra Manaus cacete!!!
        r14_13 = "0000" 'Qt registros Q18 - Zona Franca
                
        cCR14 = r14_1 & r14_2 & r14_3 & r14_4 & r14_5 & r14_6 & r14_7 & r14_8 & r14_9 & r14_10 & r14_11 & r14_12 & r14_13
        Print #iArq, cCR14
        rsCFOP_014.MoveNext
        Loop
        rsCFOP_014.Close
   Else
   cCR10 = r10_1 & r10_2 & r10_3 & r10_4 & r10_5 & r10_6 & r10_7 & r10_8 & r10_9 & r10_10 & r10_11 & r10_12
   Print #iArq, cCR10
   End If


cCount = cCount + 1
rsCFOP_ESC.MoveNext
Loop

    
'cCount = cCount + 1

'GERAR LINHAS

Call DisconnectFromDataBase

Close #iArq
'DoCmd.setwarnings (True)


End Function

