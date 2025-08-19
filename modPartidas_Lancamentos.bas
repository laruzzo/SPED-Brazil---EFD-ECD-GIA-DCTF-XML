Attribute VB_Name = "modPartidas_Lancamentos"
Option Compare Database

Public Conn As New ADODB.Connection
Public SQLStr As String

'Call ConnectToDataBase
'strSQL = ("delete from tbCadProd_Ativo_temp")
'Conn.Execute strSQL
'Call DisconnectFromDataBase

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

Public Sub Analise_CustoMedio()


Dim Db As Database
Dim rst As DAO.Recordset
Set Db = CurrentDb()

Set rst = Db.OpenRecordset("select q0.IDProd, q0.qnt, q0.CustoMedio, q0.VlrLiquido, q1.Ano, q1.Mes, q1.DataEmissao, q2.CMed_Unit, q2.DescProd from tbVendasDet as q0 inner Join tbVendas As q1 on q1.ID = q0.IDVenda inner Join tbCadProd As q2 on q0.IDProd = q2.IDProd where customedio =0 and CFOP <> 1604;")



End Sub
Public Sub CONC_BANCARIA()

'CLASSIFICAR CONCILIACAO BANCARIA
Call ConnectToDataBase

Dim Db As Database
Dim rsChaves_Conciliacao As DAO.Recordset
Set Db = CurrentDb()

'DoCmd.setwarnings (False)


Set rsChaves_Conciliacao = Db.OpenRecordset("select * from tb_Extrato_Bancario_Chaves order by Prioridade, ID")
Do Until rsChaves_Conciliacao.EOF
'DoCmd.RunSQL ("update tb_extrato_bancario set conciliacao = '" & rsChaves_Conciliacao!DESC & "' WHERE (((tb_Extrato_Bancario.Lancamento) Like '*" & rsChaves_Conciliacao!CHAVE & "*') and conciliacao is null); ")

strSQL = ("update tb_Extrato_Bancario set conciliacao = '" & rsChaves_Conciliacao!DESC & "' WHERE (((tb_Extrato_Bancario.Lancamento) Like '%" & rsChaves_Conciliacao!CHAVE & "%') and Conciliacao is null);")
Conn.Execute strSQL

rsChaves_Conciliacao.MoveNext
Loop

strSQL = ("update tb_Extrato_Bancario set conciliacao = 'Compras' where conciliacao = 'Capital Social' and Valor <0;")
Conn.Execute strSQL

Call DisconnectFromDataBase


'DoCmd.setwarnings (True)
MsgBox ("Extrato Conciliado, verifique e faça ajustes manuais necessários")
'CLASSIFICAR CONCILIACAO BANCARIA

End Sub


Public Sub PARTIDAS_DOBRADAS(cDtIni As String, cDtFim As String)

'DoCmd.setwarnings (False)
Call ConnectToDataBase


Dim Db As Database

Dim rsExtrato As DAO.Recordset

Set Db = CurrentDb()

Dim cSTR_DtINI As String
Dim cSTR_DtFIM As String

cSTR_DtINI = Replace(Format(cDtIni, "dd/mm/yyyy"), "/", "")
cSTR_DtFIM = Replace(Format(cDtFim, "dd/mm/yyyy"), "/", "")

cDtIni = Format(cDtIni, "yyyy-mm-dd")
cDtFim = Format(cDtFim, "yyyy-mm-dd")

cDtIniVb = Format(cDtIni, "mm/dd/yyyy")
cDtFimVb = Format(cDtFim, "mm/dd/yyyy")



'DoCmd.setwarnings (False)
Dim cId As Long




'LIMPA LANÇAMENTOS AUTOMATICOS E MANUAIS JÁ EXISTENTES PARA O MESMO PERÍODO

strSQL = ("delete from EFD_I200_Lancamento_Contabil where data>='" & cDtIni & "' and data <= '" & cDtFim & "';")
Conn.Execute strSQL
strSQL = ("delete from EFD_I200_Lancamento_Contabil_Head where data>='" & cDtIni & "' and data <= '" & cDtFim & "';")
Conn.Execute strSQL
'resetaAutoIncrement
strSL = ("call reset_autoincrement('EFD_I200_Lancamento_Contabil');")
Conn.Execute strSQL

strSL = ("call reset_autoincrement('EFD_I200_Lancamento_Contabil_Head');")
Conn.Execute strSQL



'LANÇA CAPITAL SOCIAL SUBSCRITO
    '80 MIL 01/01/2016
    If cDtIni = #1/1/2016# Then
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '2016-01-01' AS DATA, 80000.00 AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
        
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    
    
    'HEAD
    '2.3.1.01-Capital Social Subscrito
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '2016-01-01' AS DATA, 80000.00 AS VALOR, '" & "2.3.1.01" & "' AS CONTA, 'C' AS TIPO, 'Capital Social Subscrito conforme Contrato Social' AS HISTORICO, 'AUT', 'N', 'Contrato Social'")
    Conn.Execute strSQL
    '2.3.1.02 - Capital Social a Realizar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '2016-01-01' AS DATA, 80000.00 AS VALOR, '" & "2.3.1.02" & "' AS CONTA, 'D' AS TIPO, 'Capital Social a Realizar pelos sócios' AS HISTORICO, 'AUT','N', 'Contrato Social'")
    Conn.Execute strSQL
    '80 MIL
    Else: End If
      iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
    
'LANÇA CAPITAL SOCIAL SUBSCRITO

'LANÇA CAPITAL SOCIAL INTEGRALIZADO
Set rsExtrato = Db.OpenRecordset("Select * from tb_Extrato_Bancario WHERE data>=#" & cDtIniVb & "# and data <= #" & cDtFimVb & "# AND CONCILIACAO = 'Capital Social' ORDER BY DATA; ")

Do Until rsExtrato.EOF = True
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(Round(rsExtrato!Valor, 2), 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(Round(rsExtrato!Valor, 2), 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Entrada Caixa Capital Social conforme depósito em conta' AS HISTORICO, 'AUT','N', 'Extrato banco'")
    Conn.Execute strSQL
    '2.3.1.02 - Capital Social a Realizar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(Round(rsExtrato!Valor, 2), 2), ",", ".") & " AS VALOR, '" & "2.3.1.02" & "' AS CONTA, 'C' AS TIPO, 'Capital Social Subscrito' AS HISTORICO, 'AUT','N', 'Extrato banco'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
Loop
rsExtrato.Close
'LANÇA CAPITAL SOCIAL INTEGRALIZADO


'LANÇA SALÁRIOS
Set rsExtrato = Db.OpenRecordset("Select * from tb_Extrato_Bancario WHERE data>=#" & cDtIniVb & "# and data <= #" & cDtFimVb & "# AND CONCILIACAO = 'Salários' ORDER BY DATA; ")

Do Until rsExtrato.EOF = True
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(Round(rsExtrato!Valor, 2), 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(Round(Abs(rsExtrato!Valor), 2), 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Pagamento de salário' AS HISTORICO, 'AUT','N', 'Extrato banco'")
    Conn.Execute strSQL
    '3.3.1.01 - Despesas com salários
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(Round(Abs(rsExtrato!Valor), 2), 2), ",", ".") & " AS VALOR, '" & "3.3.1.01" & "' AS CONTA, 'D' AS TIPO, 'Pagamento de salário' AS HISTORICO, 'AUT','N', 'Extrato banco'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
Loop
rsExtrato.Close
'LANÇA SALÁRIOS


'LANÇA VENDAS REGIME DE COMPETENCIA
Set rsExtrato = Db.OpenRecordset("SELECT tbVendas.DataEmissao as Data, tbVendasDet.ValorTot + tbVendasDet.VlrDesc + tbVendasDet.Valor_IPI as Valor, tbVendasDet.Valor_ICMS, tbVendasDet.Valor_PIS, tbVendasDet.Valor_Cofins, tbVendasDet.Valor_IPI, tbVendasDet.CFOP, tbVendasDet.CFOP_DESC, tbVendas.NumNF, tbVendasDet.CustoMedio, tbVendas.ChaveNF " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIniVb & "# And (tbVendas.DataEmissao) <= #" & cDtFimVb & " 23:59:59" & "#) And ((tbVendasDet.CFOP) = '5102' Or (tbVendasDet.CFOP) = '6102' Or (tbVendasDet.CFOP) = '5101' Or (tbVendasDet.CFOP) = '6101')) " & _
"ORDER BY tbVendas.DataEmissao, tbVendas.NumNF;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    
    'RECEITA DE VENDAS
    Select Case rsExtrato!CFOP
   
    Case Is = "5101"
    '4.1.1.02    Receita de Vendas De Produtos
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "4.1.1.02" & "' AS CONTA, 'C' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
       
    Case Is = "6101"
    '4.1.1.02    Receita de Vendas De Produtos
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "4.1.1.02" & "' AS CONTA, 'C' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    Case Is = "5102"
    '4.1.1.01    Receita de Vendas De Mercadorias
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "4.1.1.01" & "' AS CONTA, 'C' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    Case Is = "6102"
    '4.1.1.01 Receita de Vendas De Mercadorias
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "4.1.1.01" & "' AS CONTA, 'C' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    End Select
    '1.1.3.01 - Clientes - Duplicatas a receber
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (ID, Data,Valor,Conta, Tipo,Historico, OPERACAO,Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.3.01" & "' AS CONTA, 'D' AS TIPO, 'Duplicatas a receber de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    
    
        'HEAD ICMS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_ICMS, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
        Conn.Execute strSQL
        Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
        rsHead.MoveLast
        cId = rsHead!ID
        'HEAD ICMS
        'ICMS
        '3.5.4 - Despesas com ICMS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_ICMS, 2), ",", ".") & " AS VALOR, '" & "3.5.4" & "' AS CONTA, 'D' AS TIPO, 'Despesas com ICMS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '2.1.1.04 - ICMS a recolher
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_ICMS, 2), ",", ".") & " AS VALOR, '" & "2.1.1.04" & "' AS CONTA, 'C' AS TIPO, 'ICMS a recolher' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
       
        'HEAD IPI
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_IPI, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
        Conn.Execute strSQL
        Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
        rsHead.MoveLast
        cId = rsHead!ID
        'HEAD IPI
         'IPI
        '3.5.5 - Despesas com IPI
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_IPI, 2), ",", ".") & " AS VALOR, '" & "3.5.5" & "' AS CONTA, 'D' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '2.1.1.03 - IPI a  recolher
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_IPI, 2), ",", ".") & " AS VALOR, '" & "2.1.1.03" & "' AS CONTA, 'C' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
         
        'HEAD PIS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_PIS, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
        Conn.Execute strSQL
        Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
        rsHead.MoveLast
        cId = rsHead!ID
        'HEAD PIS
        '3.5.6 - Despesas com PIS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_PIS, 2), ",", ".") & " AS VALOR, '" & "3.5.6" & "' AS CONTA, 'D' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '2.1.1.06 - PIS a recolher
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_PIS, 2), ",", ".") & " AS VALOR, '" & "2.1.1.06" & "' AS CONTA, 'C' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
            
        'HEAD COFINS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_Cofins, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
        Conn.Execute strSQL
        Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
        rsHead.MoveLast
        cId = rsHead!ID
        'HEAD COFINS
        '3.5.7 - Despesas com Cofins
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_Cofins, 2), ",", ".") & " AS VALOR, '" & "3.5.7" & "' AS CONTA, 'D' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '2.1.1.07 COFINS A RECOLHER
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_Cofins, 2), ",", ".") & " AS VALOR, '" & "2.1.1.07" & "' AS CONTA, 'C' AS TIPO, 'Receita de Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
  
    
    
    Select Case rsExtrato!CFOP
    Case Is = "5101", "6101"
      
    'CUSTO PRODUTO VENDIDO
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsExtrato!CustoMedio, ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '3.1.1.01   Custos dos Produtos Vendidos  -CPV
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO,Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsExtrato!CustoMedio, ",", ".") & " AS VALOR, '" & "3.1.1.01" & "' AS CONTA, 'D' AS TIPO, 'Custo Mercadoria Vendida' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    '1.1.4.02   Estoque Produtos Acabados
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO,Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsExtrato!CustoMedio, ",", ".") & " AS VALOR, '" & "1.1.4.02" & "' AS CONTA, 'C' AS TIPO, 'Custo Mercadoria Vendida' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    
       
    Case Is = "5102", "6102"
    'CUSTO MERCADORIA VENDIDA
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsExtrato!CustoMedio, ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '3.2.1.01   Custo das Mercadorias Vendidas - CMV
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data, Valor, Conta, Tipo, Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsExtrato!CustoMedio, ",", ".") & " AS VALOR, '" & "3.2.1.01" & "' AS CONTA, 'D' AS TIPO, 'Custo Mercadoria Vendida' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    '1.1.4.01   Mercadorias para revenda
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data, Valor, Conta, Tipo, Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsExtrato!CustoMedio, ",", ".") & " AS VALOR, '" & "1.1.4.01" & "' AS CONTA, 'C' AS TIPO, 'Custo Mercadoria Vendida' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    End Select
    
    
    'RECEITA DE VENDAS
    'CONTAS A RECEBER
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '1.1.3.01 - Clientes - Duplicatas a receber CREDITO - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT  " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.3.01" & "' AS CONTA, 'C' AS TIPO, 'Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT  " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'LANÇA VENDAS REGIME DE COMPETENCIA

'LANÇA AS COMPRAS REGIME DE COMPETENCIA
Set rsExtrato = Db.OpenRecordset("SELECT tbCompras.DataEmissao as Data, tbComprasDet.ValorTot + tbComprasDet.VlrFrete + tbComprasDet.VlrSeguro + tbComprasDet.VlrDesc + tbComprasDet.VlrOutro as Valor, tbComprasDet.Valor_ICMS, tbComprasDet.Valor_IPI, tbComprasDet.Valor_PIS, tbComprasDet.Valor_Cofins, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbCompras.NumNF, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= #" & cDtIniVb & "# And (tbCompras.DataEmissao) <= #" & cDtFimVb & "#) And ((tbComprasDet.CFOP_ESCRITURADA) = '1102' Or (tbComprasDet.CFOP_ESCRITURADA) = '2102' Or (tbComprasDet.CFOP_ESCRITURADA) = '1101' Or (tbComprasDet.CFOP_ESCRITURADA) = '2101')) " & _
"ORDER BY tbCompras.DataEmissao, tbCompras.NumNF;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    
    'DESPESAS DE COMPRAS
    Select Case rsExtrato!CFOP_ESCRITURADA
    
    Case Is = "1101" 'Compra para industrialização no estado
        '1.1.4.03    Estoque Insumos
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.03" & "' AS CONTA, 'D' AS TIPO, 'Compra para industrialização' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
               
    Case Is = "1102" 'Compra para comercialização no estado
        '1.1.4.01    Estoques Mercadorias para revenda
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.01" & "' AS CONTA, 'D' AS TIPO, 'Compra para revenda' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
            
    Case Is = "2102" 'Compra para comercialização fora do estado
        '1.1.4.01    Estoques Mercadorias para revenda
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id,Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.01" & "' AS CONTA, 'D' AS TIPO, 'Compra para revenda' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
    Case Is = "2101"
        '1.1.4.03    'Compra para industrialização fora do estado
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id,Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.03" & "' AS CONTA, 'D' AS TIPO, 'Compra para industrialização' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
    
    End Select
        '2.1.2.01    Fornecedores a pagar
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
        
         'HEAD ICMS
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_ICMS, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD ICMS
        'Credito ICMS - 1.1.5.02 ICMS a Recuperar
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_ICMS, 2), ",", ".") & " AS VALOR, '" & "1.1.5.02" & "' AS CONTA, 'D' AS TIPO, 'Crédito de ICMS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '1.1.4.03 Estoque Insumos ICMS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_ICMS, 2), ",", ".") & " AS VALOR, '" & "1.1.4.03" & "' AS CONTA, 'C' AS TIPO, 'Crédito de ICMS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
        
        'HEAD IPI
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_IPI, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD IPI
        'Credito IPI - 1.1.5.01 IPI a Recuperar
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_IPI, 2), ",", ".") & " AS VALOR, '" & "1.1.5.01" & "' AS CONTA, 'D' AS TIPO, 'Crédito de IPI' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '1.1.4.03 Estoque Insumos IPI
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_IPI, 2), ",", ".") & " AS VALOR, '" & "1.1.4.03" & "' AS CONTA, 'C' AS TIPO, 'Crédito de IPI' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
        Call ConnectToDataBase
        
        'HEAD PIS
         
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_PIS, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD PIS
         'Credito PIS - 1.1.5.03 PIS a Recuperar
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_PIS, 2), ",", ".") & " AS VALOR, '" & "1.1.5.03" & "' AS CONTA, 'D' AS TIPO, 'Crédito de PIS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '1.1.4.03 Estoque Insumos PIS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_PIS, 2), ",", ".") & " AS VALOR, '" & "1.1.4.03" & "' AS CONTA, 'C' AS TIPO, 'Crédito de PIS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
        
         'HEAD COFINS
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_Cofins, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD COFINS
         'Credito COFINS - 1.1.5.05 COFINS a Recuperar
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_Cofins, 2), ",", ".") & " AS VALOR, '" & "1.1.5.05" & "' AS CONTA, 'D' AS TIPO, 'Crédito de COFINS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        '1.1.4.03 Estoque Insumos COFINS
        strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico,OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor_Cofins, 2), ",", ".") & " AS VALOR, '" & "1.1.4.03" & "' AS CONTA, 'C' AS TIPO, 'Crédito de COFINS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
        Conn.Execute strSQL
        
        
    
    'DESPESA DE COMPRAS
    'CONTAS A PAGAR
     'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
   
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'LANÇA AS COMPRAS REGIME DE COMPETENCIA

'LANÇA COMPRAS COM RECIBO MANUAL
'COMPRAS DE CONSUMO COM RECIBO
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.CONSUMO, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.CONSUMO) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.4.05 Material de consumo
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.05" & "' AS CONTA, 'D' AS TIPO, 'Material de consumo' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compras material de consumo' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras material de consumo' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close

'COMPRAS DE CONSUMO COM RECIBO
'COMPRAS DE IMOBILIZADO COM RECIBO
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.IMOBILIZADO, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.IMOBILIZADO) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
  'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.2.3.03 Máquinas e Equipamentos
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.2.3.03" & "' AS CONTA, 'D' AS TIPO, 'Máquinas e equipamentos' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compras Maquinas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras Maquinas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'COMPRAS DE IMOBILIZADO COM RECIBO
'COMPRAS DE EMBALAGENS COM RECIBO
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.EMBALAGEM, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.EMBALAGEM) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD

    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.4.04 Embalagens
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.04" & "' AS CONTA, 'D' AS TIPO, 'Embalagens' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
        'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compras Embalagens' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras Embalagens' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'COMPRAS DE EMBALAGENS COM RECIBO
'COMPRAS DE MATERIAL DE ESCRITORIO COM RECIBO
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.MAT_ESCRITORIO, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.MAT_ESCRITORIO) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.4.05 Material de consumo
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.05" & "' AS CONTA, 'D' AS TIPO, 'Material de Escritório' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
        'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'COMPRAS DE  MATERIAL DE ESCRITORIO COM RECIBO
'COMPRAS DE SOFTWARE COM RECIBO
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.SOFTWARE, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.SOFTWARE) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.2.1.02 Software
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.2.1.02" & "' AS CONTA, 'D' AS TIPO, 'Aquisições de softwares' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
        'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compras Software' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras Software' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'COMPRAS DE SOFTWARE COM RECIBO
'COMPRAS DE MATERIAL PUBLICIDADE COM RECIBO
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.MAT_PUBLICIDADE, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.MAT_PUBLICIDADE) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD

    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '3.3.1.06 Despesas com Propaganda e Publicidade
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "3.3.1.06" & "' AS CONTA, 'D' AS TIPO, 'Propaganda e publicidade' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
        'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Vendas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'COMPRAS DE MATERIAL PUBLICIDADE COM RECIBO
'LANÇA COMPRAS COM RECIBO MANUAL

'TAXAS DO GOVERNO
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.TAXA_GOVERNO, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.TAXA_GOVERNO) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD

    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '3.3.1.07 Taxas do governo, licenças ambientais e outras taxas
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "3.3.1.07" & "' AS CONTA, 'D' AS TIPO, 'Taxas do governo, licenças' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Taxas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Taxas' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'TAXAS DO GOVERNO
'HONORARIOS ADVOCATICIOS, CONTABEIS E OUTROS
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.HONORARIOS, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.HONORARIOS) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '3.3.1.08    Honorários Advocaticios, Contabeis e outros honorarios
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "3.3.1.08" & "' AS CONTA, 'D' AS TIPO, 'Honorarios advocaticios ou contabeis' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'HONORARIOS ADVOCATICIOS, CONTABEIS E OUTROS
'SERVIÇOS
Set rsExtrato = Db.OpenRecordset("SELECT tbRecibo.Data, tbRecibo.Valor_Total as Valor, tbRecibo.Numero_Ref, tbCadProd.DescProd, tbCadProd.SERVICO, tbrecibo.ID_RECIBO " & _
"FROM tbCadProd INNER JOIN tbRecibo ON tbCadProd.IDProd = tbRecibo.id_Produto " & _
"WHERE (((tbRecibo.Data) >= #" & cDtIniVb & "# And (tbRecibo.Data) <= #" & cDtFimVb & "#) And ((tbCadProd.SERVICO) = 'SIM')) " & _
"ORDER BY tbRecibo.Data, tbRecibo.Numero_Ref;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Serviços contratados' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '3.3.1.09    Serviços Contratados
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "3.3.1.09" & "' AS CONTA, 'D' AS TIPO, 'Serviços contratados' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Serviços contratados' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ",  '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Serviços contratados' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!ID_RECIBO & "-" & rsExtrato!Numero_Ref & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'SERVIÇOS
'LANÇA COMPRAS COM RECIBO MANUAL


'LANÇA ALUGUEL
Set rsExtrato = Db.OpenRecordset("SELECT tbCompras.DataEmissao as Data, tbCompras.VlrTOTALNF as Valor, tbCompras.idFornecedor, tbCompras.NumNF, tbCompras.ChaveNF " & _
"FROM tbCompras " & _
"WHERE (((tbCompras.DataEmissao) >= #" & cDtIniVb & "# And (tbCompras.DataEmissao) <= #" & cDtFimVb & "#) And ((tbCompras.idFornecedor) = 1136)) " & _
"ORDER BY tbCompras.DataEmissao, tbCompras.NumNF;")
    Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '3.3.1.03 Aluguéis
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "3.3.1.03" & "' AS CONTA, 'D' AS TIPO, 'Aluguel' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Aluguel' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Aluguel' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Aluguel' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
        
   rsExtrato.MoveNext
   iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'LANÇA ALUGUEL

'LANÇA ENERGIA ELÉTRICA - Fornecedor 1131
Set rsExtrato = Db.OpenRecordset("SELECT tbCompras.DataEmissao as Data, tbCompras.VlrTOTALNF as Valor, tbCompras.ICMS_Valor,tbCompras.idFornecedor, tbCompras.NumNF, tbCompras.ChaveNF " & _
"FROM tbCompras " & _
"WHERE (((tbCompras.DataEmissao) >= #" & cDtIniVb & "# And (tbCompras.DataEmissao) <= #" & cDtFimVb & "#) And ((tbCompras.idFornecedor) = 1131)) " & _
"ORDER BY tbCompras.DataEmissao, tbCompras.NumNF;")
    Do Until rsExtrato.EOF = True
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '3.3.1.05 - Energia Elétrica
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2) - rsExtrato!ICMS_Valor, ",", ".") & " AS VALOR, '" & "3.3.1.05" & "' AS CONTA, 'D' AS TIPO, 'Conta de energia' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    'Credito ICMS - 1.1.5.02 ICMS a Recuperar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta,Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsExtrato!ICMS_Valor, ",", ".") & " AS VALOR, '" & "1.1.5.02" & "' AS CONTA, 'D' AS TIPO, 'Crédito de ICMS' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Conta de energia' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Conta de energia' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Conta de energia' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'LANÇA ENERGIA ELÉTRICA



'LANÇA RECEITA FINANCEIRA
Set rsExtrato = Db.OpenRecordset("Select * from tb_Extrato_Bancario WHERE data>=#" & cDtIniVb & "# and data <= #" & cDtFimVb & "# AND CONCILIACAO = 'Receita Financeira' ORDER BY DATA; ")
    Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '4.1.1.04    Receitas financeiras
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "4.1.1.04" & "' AS CONTA, 'C' AS TIPO, 'Receita financeira Itau conta mais' AS HISTORICO, 'AUT', 'N', 'Extrato Bancario'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Receita financeira Itau conta mais' AS HISTORICO, 'AUT', 'N', 'Extrato Bancario'")
    Conn.Execute strSQL
    
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
    Loop
rsExtrato.Close
'LANÇA RECEITA FINANCEIRA

'TARIFAS BANCÁRIAS
Set rsExtrato = Db.OpenRecordset("Select * from tb_Extrato_Bancario WHERE data>=#" & cDtIniVb & "# and data <= #" & cDtFimVb & "# AND CONCILIACAO = 'Tarifa Bancária' ORDER BY DATA; ")
    Do Until rsExtrato.EOF = True
 'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2) * -1, ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Receita financeira Itau conta mais' AS HISTORICO, 'AUT', 'N', 'Extrato Bancario'")
    Conn.Execute strSQL
    
    '3.3.1.04 - Tarifas Bancárias
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2) * -1, ",", ".") & " AS VALOR, '" & "3.3.1.04" & "' AS CONTA, 'D' AS TIPO, 'Receita financeira Itau conta mais' AS HISTORICO, 'AUT', 'N', 'Extrato Bancario'")
    Conn.Execute strSQL
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
    End If
    Loop
rsExtrato.Close
'TARIFAS BANCÁRIAS

'IMOBILIZADO
Set rsExtrato = Db.OpenRecordset("SELECT tbCompras.DataEmissao as Data, tbCompras.VlrTOTALNF as Valor, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbCompras.NumNF, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= #" & cDtIniVb & "# And (tbCompras.DataEmissao) <= #" & cDtFimVb & "#) And ((tbComprasDet.CFOP_ESCRITURADA) = '1551' Or (tbComprasDet.CFOP_ESCRITURADA) = '2551')) " & _
"ORDER BY tbCompras.DataEmissao, tbCompras.NumNF;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '1.2.3.03    Ativo Imobilizado - Maquinas e Equipamentos
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.2.3.03" & "' AS CONTA, 'D' AS TIPO, 'Compras Imobilizado' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras Imobilizado' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compras Imobilizado' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compras Imobilizado' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
    End If
Loop
rsExtrato.Close

'IMOBILIZADO

'DEPRECIAÇÃO DO IMOBILIZADO
'1.2.3.98    (-) Depreciação Acumulada
Set rsImobCad = Db.OpenRecordset("SELECT * FROM tbImobilizado where ciclo = '1' order by DataEmissao Asc")
Do Until rsImobCad.EOF
    
    Set rsExtrato = Db.OpenRecordset("select Ciclo, ChaveNfe, DataEmissao, IDProd, Sum(ValorTot) as ValorTotAtual from tbImobilizado where IDProd = " & rsImobCad!IDProd & " and DataEmissao >= #" & cDtIniVb & "# and DataEmissao <= #" & cDtFimVb & "# and int(ciclo) > 1 and chaveNfe = '" & rsImobCad!chaveNFE & "' group by Ciclo, ChaveNfe, DataEmissao, IDProd order by int(Ciclo) asc;")
    If rsExtrato.EOF = True And rsExtrato.BOF = True Then
    GoTo pulaDepre:
    Else
    End If
    
    
    Do Until rsExtrato.EOF = True
    
    Set rsExtratoAnt = Db.OpenRecordset("select Ciclo, ChaveNfe, DataEmissao, IDProd, Sum(ValorTot) as ValorTotAtual from tbImobilizado where IDProd = " & rsImobCad!IDProd & " and int(ciclo) = " & Int(rsExtrato!Ciclo) - 1 & " and chaveNfe = '" & rsExtrato!chaveNFE & "' group by Ciclo, ChaveNfe, DataEmissao, IDProd;")
    If rsExtratoAnt.EOF = True And rsExtratoAnt.BOF = True Then
    GoTo pulaproximo:
    Else
    End If
    
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!DataEmissao, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtratoAnt!ValorTotAtual - rsExtrato!ValorTotAtual, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '1.2.3.98   (-) Depreciação Acumulada
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!DataEmissao, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtratoAnt!ValorTotAtual - rsExtrato!ValorTotAtual, 2), ",", ".") & " AS VALOR, '" & "1.2.3.98" & "' AS CONTA, 'C' AS TIPO, 'Depreciação do ativo imobilizado - IDProd: " & rsExtrato!IDProd & " Ciclo: " & rsExtrato!Ciclo & "' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chaveNFE & "'")
    Conn.Execute strSQL
    
    '3.3.1.10 Despesa de Depreciação
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!DataEmissao, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtratoAnt!ValorTotAtual - rsExtrato!ValorTotAtual, 2), ",", ".") & " AS VALOR, '" & "3.3.1.10" & "' AS CONTA, 'D' AS TIPO, 'Depreciação do ativo imobilizado - IDProd: " & rsExtrato!IDProd & " Ciclo: " & rsExtrato!Ciclo & "' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chaveNFE & "'")
    Conn.Execute strSQL
    
    

pulaproximo:
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
    End If
    Loop
pulaDepre:
    rsExtrato.Close
    rsImobCad.MoveNext
Loop
'1.2.3.99    (-) Amortização Acumulada
'DEPRECIAÇÃO DO IMOBILIZADO

'COMPRA DE CONSUMO 1.1.4.05 Material de consumo
Set rsExtrato = Db.OpenRecordset("SELECT tbCompras.DataEmissao as Data, tbCompras.VlrTOTALNF as Valor, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbCompras.NumNF, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= #" & cDtIniVb & "# And (tbCompras.DataEmissao) <= #" & cDtFimVb & "#) And ((tbComprasDet.CFOP_ESCRITURADA) = '2556' Or (tbComprasDet.CFOP_ESCRITURADA) = '1556')) " & _
"ORDER BY tbCompras.DataEmissao, tbCompras.NumNF;")

Do Until rsExtrato.EOF = True
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '1.1.4.05    Estoque Material de consumo
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.4.05" & "' AS CONTA, 'D' AS TIPO, 'Compra para uso e consumo' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
'    Call ConnectToDataBase
    Conn.Execute strSQL
    '2.1.2.01    Fornecedores a pagar
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compra para uso e consumo' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
   'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
    Conn.Execute strSQL
    
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.1.2.01 - Fornecedores A PAGAR - QUANDO PAGOU
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "2.1.2.01" & "' AS CONTA, 'D' AS TIPO, 'Compra para uso e consumo' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    '1.1.2.01 - Banco Itau AG 4522 CT 25055-6
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & "1.1.2.01" & "' AS CONTA, 'C' AS TIPO, 'Compra para uso e consumo' AS HISTORICO, 'AUT', 'N', '" & rsExtrato!chavenf & "'")
    Conn.Execute strSQL
    
    
    rsExtrato.MoveNext
    iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If
Loop
rsExtrato.Close
'COMPRA DE CONSUMO 1.1.4.05    Material de consumo

'LANCAMENTO DE CREDITOS DE IMPOSTOS APROPRIADOS NO MÊS - CIAP
'APENAS CIAP, PORQUE OS OUTROS CREDITOS JÁ FORAM CONSIDERADOS NO CUSTO DO PRODUTO VENDIDO E O SALDO DE IMPOSTOS E DO ATIVO CIRCULANTE.
'MAS ATENÇÃO SÓ LANÇA O CIAP SE O DEBITO DE ICMS FOR MAIOR. SENÃO A CONTA FICA NEGATIVA.

Dim valCIAP As Double
Dim cDataCiap As Date
valCIAP = 0
Set rsExtrato = Db.OpenRecordset("SELECT * FROM tbResumo_ICMS " & _
"WHERE ANO = '" & year(cDtIniVb) & "' And MES  = " & month(Format(cDtIniVb, "mm/dd/yyyy")) & " ;")
    
    Do Until rsExtrato.EOF
        
        cDataCiap = Format(str(rsExtrato!MES) + "/" + "01" + "/" + str(rsExtrato!ANO), "mm/dd/yyyy")
        
        'HEAD ICMS CIAP
         'If rsExtrato!DEB >= valCIAP Then
         If rsExtrato!DEB >= rsExtrato!CIAP Then
         valCIAP = rsExtrato!CIAP
         Else
         valCIAP = rsExtrato!DEB
         End If
         
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(cDataCiap, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(valCIAP, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD ICMS
        'ICMS
        '3.5.4 - Despesas com ICMS
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDataCiap, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(valCIAP, 2), ",", ".") & " AS VALOR, '" & "3.5.4" & "' AS CONTA, 'C' AS TIPO, 'Crédido de CIAP ICMS' AS HISTORICO, 'AUT', 'N', 'CIAP'")
         Conn.Execute strSQL
         '2.1.1.04 - ICMS a recolher
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDataCiap, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(valCIAP, 2), ",", ".") & " AS VALOR, '" & "2.1.1.04" & "' AS CONTA, 'D' AS TIPO, 'ICMS a recuperar CIAP' AS HISTORICO, 'AUT', 'N', 'CIAP'")
         Conn.Execute strSQL
        
        
    rsExtrato.MoveNext
    Loop
       
'LANCAMENTO DE CREDITOS DE IMPOSTOS APROPRIADOS NO MÊS - CIAP

'LANÇAMENTO DE IRPJ E CSSL APENAS PARA LUCRO PRESUMIDO
Set rsExtrato = Db.OpenRecordset("SELECT * FROM tbResumo_IRPJ_CSLL " & _
"WHERE ANO = '" & year(cDtIniVb) & "' And MES  = " & month(Format(cDtIniVb, "mm/dd/yyy")) & " AND REGIME = 'PRESUMIDO';")

 Do Until rsExtrato.EOF
        
        cDataIR = Format(str(rsExtrato!MES) + "/" + "31" + "/" + str(rsExtrato!ANO), "mm/dd/yyyy")
        
        'HEAD IRPJ
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(cDataIR, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!VL_IRPJ, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD IRPJ
        'IRPJ
        '3.6.2 - Despesas com IRPJ
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDataIR, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!VL_IRPJ, 2), ",", ".") & " AS VALOR, '" & "3.6.2" & "' AS CONTA, 'D' AS TIPO, 'Provisão para Imposto de Renda - IRPJ' AS HISTORICO, 'AUT', 'N', 'IRPJ'")
         Conn.Execute strSQL
         '2.1.1.10 - IRPJ a Recolher
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDataIR, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!VL_IRPJ, 2), ",", ".") & " AS VALOR, '" & "2.1.1.10" & "' AS CONTA, 'C' AS TIPO, 'IRPJ a Recolher' AS HISTORICO, 'AUT', 'N',                     'IRPJ'")
         Conn.Execute strSQL


         'HEAD CSLL
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(cDataIR, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!VL_CSLL, 2), ",", ".") & " AS VALOR, 'AUT', 'N'")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD CSLL
        'CSLL
        '3.6.1 - Despesas com CSLL
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDataIR, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!VL_CSLL, 2), ",", ".") & " AS VALOR, '" & "3.6.1" & "' AS CONTA, 'D' AS TIPO, 'Provisão para Imposto de Renda - CSLL' AS HISTORICO, 'AUT', 'N', 'CSLL'")
         Conn.Execute strSQL
         '2.1.1.09 - CSLL a Recolher
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDataIR, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!VL_CSLL, 2), ",", ".") & " AS VALOR, '" & "2.1.1.09" & "' AS CONTA, 'C' AS TIPO, 'CSLL a Recolher' AS HISTORICO, 'AUT', 'N', 'CSLL'")
         Conn.Execute strSQL
    
         
         'FALTA LANÇAR NAS CONTAS PATRIMONIAIS D NA CONTA PASSIVO E C NA CONTA CAIXA CORRENTE
     
    rsExtrato.MoveNext
    Loop
'LANÇAMENTO DE IRPJ E CSSL


lancamento:
Dim cAnome As String
cAnomes = year(cDtIniVb) & Format((month(Format(cDtIniVb, "mm/dd/yyyy"))), "00")


'PROCESSA LANÇAMENTOS MANUAIS ADICIONAIS
Set rsExtrato = Db.OpenRecordset("SELECT * FROM efd_i200_lancamento_manual_head " & _
"WHERE ANOMES = '" & cAnomes & "';")


    
    Do Until rsExtrato.EOF
        
        ID = rsExtrato!ID
        
        'HEAD Lançamento Manual
        
         
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(rsExtrato!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtrato!Valor, 2), ",", ".") & " AS VALOR, '" & rsExtrato!Operacao & "' as Operacao,'" & rsExtrato!Indicador & "' as indicador;")
         Conn.Execute strSQL
         Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
         rsHead.MoveLast
         cId = rsHead!ID
        'HEAD Lançamento Manual
        
         Set rsExtratoDet = Db.OpenRecordset("SELECT * FROM efd_i200_lancamento_manual " & _
         "WHERE ID_Head = " & rsExtrato!ID & ";")
         Do Until rsExtratoDet.EOF
         strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(rsExtratoDet!Data, "yyyy-mm-dd") & "' AS DATA, " & Replace(Round(rsExtratoDet!Valor, 2), ",", ".") & " AS VALOR, '" & rsExtratoDet!Conta & "' AS CONTA, '" & rsExtratoDet!Tipo & "' AS TIPO, '" & rsExtratoDet!Historico & "' as Historico, '" & rsExtratoDet!Operacao & "' as Operacao, '" & rsExtratoDet!Indicador & "' as indicador, '" & rsExtratoDet!Num_Nota & "' as numnota")
         Conn.Execute strSQL
         rsExtratoDet.MoveNext
         Loop
        
        
    rsExtrato.MoveNext
    Loop


'PROCESSA LANÇAMENTOS MANUAIS ADICIONAIS


'TIRA NEGATIVOS DO HEAD
strSQL = ("UPDATE EFD_I200_Lancamento_Contabil_Head SET EFD_I200_Lancamento_Contabil_Head.Valor = EFD_I200_Lancamento_Contabil_Head.Valor*-1 WHERE (((EFD_I200_Lancamento_Contabil_Head.Valor)<0)); ")
Conn.Execute strSQL



'ADICIONA ANOMES NOS LANÇAMENTOS
strSQL = ("UPDATE EFD_I200_Lancamento_Contabil " & _
"Set EFD_I200_Lancamento_Contabil.AnoMes = concat(Year(EFD_I200_Lancamento_Contabil.Data), Lpad(Month(EFD_I200_Lancamento_Contabil.Data), 2, 0)) " & _
"WHERE EFD_I200_Lancamento_Contabil.AnoMes is null;")
Conn.Execute strSQL




iCounter = iCounter + 1
      If iCounter = 50 Then
      DoEvents
      iCounter = 0
      End If

'DoCmd.setwarnings (True)


End Sub



