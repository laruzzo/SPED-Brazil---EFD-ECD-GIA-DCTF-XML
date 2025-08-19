Attribute VB_Name = "modSaldos_Periodicos"
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

Public Sub Contabilizacao_Saldos_Periodicos(cDtIni As String, cDtFim As String)

'DoCmd.setwarnings (False)
Call ConnectToDataBase


Dim Db As Database
Set Db = CurrentDb()



Dim cAno As Integer
Dim cMes As Integer
Dim cAnomes As String

Dim cAnoFIm As Integer
Dim cMesFim As Integer
Dim cAnoMesFim As String
Dim cDtINICIO As String
Dim cDtFinal As String



cAno = year(Format(cDtIni, "mm/dd/yyyy"))
cMes = Left(Format(cDtIni, "mm/dd/yyyy"), 2)

cAnomes = cAno & Format(cMes, "00")

cAnoFIm = year(Format(cDtFim, "mm/dd/yyyy"))
cMesFim = Left(Format(cDtFim, "mm/dd/yyyy"), 2)
cAnoMesFim = cAnoFIm & Format(cMesFim, "00")



'DoCmd.RunSQL ("delete * from EFD_I150_Calendario where anomes >=" & cAnoMes & " and anomes <=" & cAnoMesFim & "")

'Do Until cAnoMes = cAnoMesFim
'cAnoMes = cAno & Format(cMes, "00")
'cAnoMesFim = cAnoFIm & Format(cMesFim, "00")
'cDtINICIO = DateSerial(cAno, cMes, 1)
'cDtINICIO = Format(cDtINICIO, "mm/dd/yyyy")
'If cMes = 12 Then
'cDtFinal = DateAdd("d", -1, DateSerial(cAno + 1, 1, 1))
'cDtFinal = Format(cDtFinal, "mm/dd/yyyy")
'Else
'cDtFinal = DateAdd("d", -1, DateSerial(cAno, cMes + 1, 1))
'cDtFinal = Format(cDtFinal, "mm/dd/yyyy")
'End If

'DoCmd.RunSQL ("insert into EFD_I150_Calendario (ANOMES, ANO, MES, DT_INI, DT_FIM) select " & cAnoMes & ", " & cAno & ", " & cMes & ", #" & cDtINICIO & "#, #" & cDtFinal & "#")
'cMes = cMes + 1
'If cMes = 13 Then
'cMes = 1
'cAno = cAno + 1
'Else
'End If
'Loop


'DoCmd.Rename "EFD_I150_Calendario_Temp", acTable, "EFD_I150_Calendario"
'DoCmd.RunSQL ("SELECT EFD_I150_Calendario_Temp.ANOMES, EFD_I150_Calendario_Temp.ANO, EFD_I150_Calendario_Temp.MES, EFD_I150_Calendario_Temp.DT_INI, EFD_I150_Calendario_Temp.DT_FIM INTO EFD_I150_Calendario " & _
'"FROM EFD_I150_Calendario_Temp " & _
'"ORDER BY EFD_I150_Calendario_Temp.ANOMES;")
'DoCmd.DeleteObject acTable, "EFD_I150_Calendario_Temp"

  


Dim cSaldoIni As Double
Dim cSaldoFim As Double

cSaldoIni = 0
cSaldoFim = 0


  Set rsI150 = Db.OpenRecordset("SELECT * FROM EFD_I150_Calendario where anomes >=" & cAnomes & " and anomes <=" & cAnoMesFim & "")
  
     If rsI150!ENCERRAMENTO = "S" Then
        'DELETA LANÇAMENTOS DE RESULTADO JÁ EFETUADOS PARA EVITAR DUPLICIDADE
        strSQL = ("DELETE FROM EFD_I200_Lancamento_Contabil where Indicador = 'E' and anomes = " & cAnomes & " ")
        Conn.Execute strSQL
     Else
     End If
    
    
   Dim cSaldoD As Double
   Dim cSaldoC As Double
   
  Do Until rsI150.EOF
  strSQL = ("delete from EFD_I155_Detalhe_Saldos where anomes = " & rsI150!ANOMES & "")
  Conn.Execute strSQL
  
    
  'DEBITOS E CREDITOS
  Set rsSaldos = Db.OpenRecordset("TRANSFORM Sum(EFD_I200_Lancamento_Contabil.Valor) as Valor " & _
    "SELECT EFD_I200_Lancamento_Contabil.Conta " & _
    "FROM EFD_I200_Lancamento_Contabil " & _
    "where EFD_I200_Lancamento_Contabil.Data >= #" & Format(rsI150!DT_INI, "mm/dd/yyyy") & "# and EFD_I200_Lancamento_Contabil.Data <= #" & Format(rsI150!DT_FIM, "mm/dd/yyyy") & "# " & _
    "GROUP BY EFD_I200_Lancamento_Contabil.Conta " & _
    "PIVOT EFD_I200_Lancamento_Contabil.Tipo;")
    
    Do Until rsSaldos.EOF
    cSaldoD = 0
    cSaldoC = 0
    
    If IsNull(rsSaldos.Fields("D").Value) Then
    Else
    cSaldoD = rsSaldos!D
    End If
    If IsNull(rsSaldos.Fields("C").Value) Then
    Else
    cSaldoC = rsSaldos!c
    End If
    strSQL = ("INSERT INTO EFD_I155_Detalhe_Saldos (CHAVE, ANOMES, Data_INI, Data_FIM, Cod_Conta, Total_Debitos, Total_Creditos) " & _
    "SELECT '" & rsI150!ANOMES & rsSaldos!Conta & "' as CHAVE, '" & rsI150!ANOMES & "' as ANOMES, " & _
    "'" & Format(rsI150!DT_INI, "yyyy-mm-dd") & "' as DTINI, " & _
    "'" & Format(rsI150!DT_FIM, "yyyy-mm-dd") & "' as DTFIM, " & _
    "'" & rsSaldos!Conta & "' as CONTA, " & _
    "" & Replace(cSaldoD, ",", ".") & " as D, " & _
    "" & Replace(cSaldoC, ",", ".") & " as C;")
    Conn.Execute strSQL
    
    rsSaldos.MoveNext
    iCounter = iCounter + 1
      If iCounter = 100 Then
      DoEvents
      iCounter = 0
      End If
    Loop
    
    'DEBITOS E CREDITOS DE CONTAS COM SALDO DO PERÍODO ANTERIOR SEM MOVIMENTO NESSE PERIODO
    Set rsSaldos = Db.OpenRecordset("TRANSFORM Sum(EFD_I200_Lancamento_Contabil.Valor) AS Valor " & _
        "SELECT EFD_I200_Lancamento_Contabil.Conta " & _
        "FROM EFD_I200_Lancamento_Contabil " & _
        "WHERE (((EFD_I200_Lancamento_Contabil.Data) < #" & Format(rsI150!DT_INI, "mm/dd/yyyy") & "#)) " & _
        "GROUP BY EFD_I200_Lancamento_Contabil.Conta " & _
        "PIVOT EFD_I200_Lancamento_Contabil.Tipo;")
    
    Do Until rsSaldos.EOF
    cSaldoD = 0
    cSaldoC = 0
    If IsNull(rsSaldos.Fields("D").Value) Then
    Else
    cSaldoD = rsSaldos!D
    End If
    If IsNull(rsSaldos.Fields("C").Value) Then
    Else
    cSaldoC = rsSaldos!c
    End If
    
    cSaldoIni = cSaldoD + cSaldoC
    cSaldoFim = cSaldoIni
    
    Set rsTest = Db.OpenRecordset("select CHAVE from EFD_I155_Detalhe_Saldos where CHAVE = '" & rsI150!ANOMES & rsSaldos!Conta & "'")
    If rsTest.RecordCount = 0 Then
    
    strSQL = ("INSERT INTO EFD_I155_Detalhe_Saldos (CHAVE, ANOMES, Data_INI, Data_FIM, Cod_Conta, Saldo_Inicial, Saldo_Final) " & _
    "SELECT '" & rsI150!ANOMES & rsSaldos!Conta & "' as CHAVE, '" & rsI150!ANOMES & "' as ANOMES, " & _
    "'" & Format(rsI150!DT_INI, "yyyy-mm-dd") & "' as DTINI, " & _
    "'" & Format(rsI150!DT_FIM, "yyyy-mm-dd") & "' as DTFIM, " & _
    "'" & rsSaldos!Conta & "' as CONTA, " & _
    "" & Replace(cSaldoIni, ",", ".") & " as D, " & _
    "" & Replace(cSaldoFim, ",", ".") & " as C;")
    Conn.Execute strSQL
    
    Else
    End If
    
    rsSaldos.MoveNext
    iCounter = iCounter + 1
      If iCounter = 100 Then
      DoEvents
      iCounter = 0
      End If
    Loop
    
   
    
     'TESTA SE O MÊS TEM ENCERRAMENTO E LANÇA
    If rsI150!ENCERRAMENTO = "S" Then
    'Call Refresh_Saldos(rsI150!ANOMES) - o Refresh está dentro do ARE
    Call Apuracao_Resultado_ARE(rsI150!ANOMES, rsI150!DT_FIM)
    Call Contabilizacao_Saldos_Periodicos_Fechamento(cAnomes)
   'Call Refresh_Saldos_Fechamento(cAnomes)
   'Call Resultado_Exercicio_Anterior(cAnomes)
    Else
    End If
    
    Call Refresh_Saldos(rsI150!ANOMES)
        
rsI150.MoveNext
iCounter = iCounter + 1
      If iCounter = 100 Then
      DoEvents
      iCounter = 0
      End If
Loop

  
  
iCounter = iCounter + 1
      If iCounter = 100 Then
      DoEvents
      iCounter = 0
      End If



'LIMPA ZERO ZERO ZERO SALDO
Call ConnectToDataBase
strSQL = ("DELETE FROM EFD_I155_Detalhe_Saldos " & _
"WHERE (((EFD_I155_Detalhe_Saldos.Saldo_Inicial)=0) AND ((EFD_I155_Detalhe_Saldos.Total_Debitos)=0) AND ((EFD_I155_Detalhe_Saldos.Total_Creditos)=0) AND ((EFD_I155_Detalhe_Saldos.Saldo_Final)=0));")
Conn.Execute strSQL

Call DisconnectFromDataBase


End Sub

Private Sub Refresh_Saldos(cAnomes As String)

Call ConnectToDataBase

Dim Db As Database
Set Db = CurrentDb()

Set rsContas = Db.OpenRecordset("SELECT COD_CONTA FROM EFD_I155_Detalhe_Saldos where anomes = " & cAnomes & " group by COD_Conta")

strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Saldo_Inicial = 0, Saldo_Final = 0, Ind_Saldo_INI = 0, Ind_Saldo_FIM = 0, Diferenca = 0  where anomes = " & cAnomes & "")
Conn.Execute strSQL

'strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Saldo_Inicial = 0, Saldo_Final = 0, Ind_Saldo_INI = 0, Ind_Saldo_FIM = 0, Diferenca = 0 where Saldo_Inicial is null ;")

strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Total_Creditos = 0 where Total_Creditos is Null and anomes = " & cAnomes & "")
Conn.Execute strSQL
strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Total_Debitos = 0 where Total_Debitos is Null and anomes = " & cAnomes & "")
Conn.Execute strSQL



'strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Total_Creditos = 0 = 0 , Total_Debitos = 0 where Total_Creditos is null;")
'Conn.Execute strSQL
'Para fechamento Mensal fazer assim, para fechamento anual fazer cAnomesAnt = cAno = cAno-1
cAno = Left(cAnomes, 4)
cMes = Right(cAnomes, 2)
'cAno = cAno - 1
cMesAnt = Int(cMes) - 1
If Int(cMesAnt) = 0 Then
cMesAnt = 12
cAnoAnt = cAno - 1
cAnomesAnt = cAnoAnt & Format(cMesAnt, "00")
Else
cAnomesAnt = cAno & Format(cMesAnt, "00")
End If


Do Until rsContas.EOF = True
    cSaldoIni = 0
    cSaldoFim = 0
    
    Set rsContasHist = Db.OpenRecordset("SELECT * FROM EFD_I155_Detalhe_Saldos WHERE COD_CONTA = '" & rsContas!Cod_Conta & "' and anomes = " & cAnomesAnt & "")
    Set rsContasAtual = Db.OpenRecordset("SELECT * FROM EFD_I155_Detalhe_Saldos WHERE COD_CONTA = '" & rsContas!Cod_Conta & "' and anomes = " & cAnomes & "")
    If rsContasHist.RecordCount = 0 Then
    cSaldoIni = 0
    Else
    cSaldoIni = rsContasHist!Saldo_Final
    End If
    
    cSaldoFim = cSaldoIni + rsContasAtual!Total_Creditos - rsContasAtual!Total_Debitos
    
    strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Saldo_Inicial = " & Replace(cSaldoIni, ",", ".") & ", Saldo_Final = " & Replace(cSaldoFim, ",", ".") & ", Diferenca = " & Replace(cSaldoIni - cSaldoFim, ",", ".") & " where CHAVE = '" & cAnomes & rsContas!Cod_Conta & "'")
    Conn.Execute strSQL
    'DoCmd.RunSQL ("UPDATE EFD_I155_Detalhe_Saldos_Antes_Encerramento SET Saldo_Inicial = " & Replace(cSaldoIni, ",", ".") & ", Saldo_Final = " & Replace(cSaldoFim, ",", ".") & ", Diferenca = " & Replace(cSaldoIni - cSaldoFim, ",", ".") & " where CHAVE = '" & rsContasHist!ANOMES & rsContasHist!Cod_Conta & "'")
    
    rsContas.MoveNext
    
Loop

strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Ini = 'C' where saldo_Inicial >= 0 and anomes> " & cAnomesAnt & " and anomes <= " & cAnomes & "  ")
Conn.Execute strSQL
strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Ini = 'D' where saldo_Inicial < 0 and anomes> " & cAnomesAnt & " and anomes <= " & cAnomes & "")
Conn.Execute strSQL
strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Fim = 'C' where saldo_Final >= 0 and anomes> " & cAnomesAnt & " and anomes <= " & cAnomes & " ")
Conn.Execute strSQL
strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Fim = 'D' where saldo_Final < 0 and anomes> " & cAnomesAnt & " and anomes <= " & cAnomes & " ")
Conn.Execute strSQL

Call DisconnectFromDataBase

End Sub
Private Sub Apuracao_Resultado_ARE(cAnomes, cDtFim)

Call ConnectToDataBase

''DoCmd.setwarnings (False)
cDtFim = Format(cDtFim, "mm/dd/yyyy")
Dim Db As Database
Set Db = CurrentDb()

'DELETA LANÇAMENTOS DE RESULTADO JÁ EFETUADOS PARA EVITAR DUPLICIDADE
strSQL = ("DELETE FROM EFD_I200_Lancamento_Contabil where Indicador = 'E' and anomes = " & cAnomes & " ")
Conn.Execute strSQL
'precisa deletar o head também. o head deleta automatico por causa da chave
strSQL = ("DELETE FROM overturecervej01.EFD_I350_Detalhe_Saldos_Antes_Encerramento WHERE ANOMES = " & cAnomes & " ")
Conn.Execute strSQL

Call Refresh_Saldos(RTrim(LTrim(str(cAnomes))))

Set rsSaldos = Db.OpenRecordset("SELECT ANOMES, q1.Total_Despesas, q1.Total_Receitas, Total_Receitas-Total_Despesas AS Apuracao " & _
"FROM (SELECT ANOMES, Sum(EFD_I155_Detalhe_Saldos.Total_Debitos) AS Total_Despesas, Sum(EFD_I155_Detalhe_Saldos.Total_Creditos) AS Total_Receitas " & _
"FROM EFD_I155_Detalhe_Saldos INNER JOIN EFD_I050_Contas_Contabeis ON EFD_I155_Detalhe_Saldos.Cod_Conta = EFD_I050_Contas_Contabeis.COD " & _
"WHERE (Left(Cod_Conta, 1) = 3 Or Left(Cod_Conta, 1) = 4) And EFD_I050_Contas_Contabeis.TIPO = 'A' " & _
"GROUP by ANOMES)  AS q1 " & _
"WHERE q1.ANOMES=" & cAnomes & ";")

Dim cValorC As Double
Dim cValorD As Double
Dim cValorT As Double
cValorC = 0
cValorD = 0
cValorT = 0

Call ConnectToDataBase
Conn.Execute strSQL
    
    
Do Until rsSaldos.EOF

   'Lançamento Prejuizo
        If rsSaldos!APURACAO < 0 Then
    'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsSaldos!APURACAO * -1, ",", ".") & " AS VALOR, 'AUT', 'E'")
    Conn.Execute strSQL
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.3.3.03 Prejuízos Acumulados D
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsSaldos!APURACAO * -1, ",", ".") & " AS VALOR, '" & "2.3.3.03" & "' AS CONTA, 'D' AS TIPO, 'Prejuízos Acumulados no Período - Resultado' AS HISTORICO, 'AUT', 'E', 'Resultado'")
    Conn.Execute strSQL
    'Encerramento de Contas de Depesas como C
    Set rsResultado = Db.OpenRecordset("SELECT EFD_I155_Detalhe_Saldos.CHAVE, EFD_I155_Detalhe_Saldos.ANOMES, EFD_I155_Detalhe_Saldos.Cod_Conta, EFD_I050_Contas_Contabeis.DESCRICAO, EFD_I050_Contas_Contabeis.NATUREZA, EFD_I155_Detalhe_Saldos.Saldo_Final, EFD_I155_Detalhe_Saldos.Ind_Saldo_Fim " & _
    "FROM EFD_I050_Contas_Contabeis INNER JOIN EFD_I155_Detalhe_Saldos ON EFD_I050_Contas_Contabeis.COD = EFD_I155_Detalhe_Saldos.Cod_Conta " & _
    "WHERE (((EFD_I155_Detalhe_Saldos.ANOMES)=" & cAnomes & ") AND (Left(Cod_Conta,1)='3') AND ((EFD_I050_Contas_Contabeis.NATUREZA)='04'));")
    Do Until rsResultado.EOF
    'Manda pra Contas antes do resultado
    strSQL = ("DELETE FROM EFD_I350_Detalhe_Saldos_Antes_Encerramento where CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    strSQL = ("INSERT INTO EFD_I350_Detalhe_Saldos_Antes_Encerramento (CHAVE,ANOMES,Data_INI,Data_FIM,Cod_Conta,Ind_Saldo_Ini,Saldo_Inicial,Total_Debitos,Total_Creditos,Saldo_Final,Ind_Saldo_Fim,Diferenca) select * FROM EFD_I155_Detalhe_Saldos WHERE  CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    'Faz o Lançamento Zerando
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsResultado!Saldo_Final * -1, ",", ".") & " AS VALOR, '" & rsResultado!Cod_Conta & "' AS CONTA, 'C' AS TIPO, 'Encerramento de contas de Receitas' AS HISTORICO, 'AUT', 'E', 'Resultado'")
    Conn.Execute strSQL
    
    cValorC = cValorC + (rsResultado!Saldo_Final * -1)
    
    rsResultado.MoveNext
    Loop
    
    'Encerramento de Contas de Receitas como D
    Set rsResultado = Db.OpenRecordset("SELECT EFD_I155_Detalhe_Saldos.CHAVE, EFD_I155_Detalhe_Saldos.ANOMES, EFD_I155_Detalhe_Saldos.Cod_Conta, EFD_I050_Contas_Contabeis.DESCRICAO, EFD_I050_Contas_Contabeis.NATUREZA, EFD_I155_Detalhe_Saldos.Saldo_Final, EFD_I155_Detalhe_Saldos.Ind_Saldo_Fim " & _
    "FROM EFD_I050_Contas_Contabeis INNER JOIN EFD_I155_Detalhe_Saldos ON EFD_I050_Contas_Contabeis.COD = EFD_I155_Detalhe_Saldos.Cod_Conta " & _
    "WHERE (((EFD_I155_Detalhe_Saldos.ANOMES)=" & cAnomes & ") AND (Left(Cod_Conta,1)='4') AND ((EFD_I050_Contas_Contabeis.NATUREZA)='04'));")
    Do Until rsResultado.EOF
    'Manda pra Contas antes do resultado
    strSQL = ("DELETE FROM EFD_I350_Detalhe_Saldos_Antes_Encerramento where CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    strSQL = ("INSERT INTO EFD_I350_Detalhe_Saldos_Antes_Encerramento (CHAVE,ANOMES,Data_INI,Data_FIM,Cod_Conta,Ind_Saldo_Ini,Saldo_Inicial,Total_Debitos,Total_Creditos,Saldo_Final,Ind_Saldo_Fim,Diferenca) SELECT * FROM EFD_I155_Detalhe_Saldos WHERE  CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    'Faz o Lançamento Zerando
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsResultado!Saldo_Final, ",", ".") & " AS VALOR, '" & rsResultado!Cod_Conta & "' AS CONTA, 'D' AS TIPO, 'Encerramento de contas de Despesas' AS HISTORICO, 'AUT', 'E', 'Resultado'")
    Conn.Execute strSQL
       
    cValorD = cValorD + (rsResultado!Saldo_Final)
    
    
    rsResultado.MoveNext
    Loop
    'atualiza o head
    strSQL = ("UPDATE EFD_I200_Lancamento_Contabil_Head SET Valor = " & Replace(Round(cValorC, 2), ",", ".") & " where EFD_I200_Lancamento_Contabil_Head.ID = " & cId & " ")
    Conn.Execute strSQL
    cValorD = 0
    cValorC = 0
    cValorT = 0
    'atualiza o head
    
    Else
    End If
    
    'Lançamento Lucro
    If rsSaldos!APURACAO > 0 Then

    'ENCERRAMENTO CONTA LUCROS ACUMULADOS TRANSFERENCIA PARA RESERVAS DE LUCROS
    'https://www.ebah.com.br/content/ABAAABDGYAF/contas-resultado
     'HEAD
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsSaldos!APURACAO, ",", ".") & " AS VALOR, 'AUT', 'E'")
    Conn.Execute strSQL
    Set rsHead = Db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
    rsHead.MoveLast
    cId = rsHead!ID
    'HEAD
    '2.3.3.02 Lucros Acumulados C
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsSaldos!APURACAO, ",", ".") & " AS VALOR, '" & "2.3.3.02" & "' AS CONTA, 'C' AS TIPO, 'Lucros Acumulados no Período - Resultado' AS HISTORICO, 'AUT', 'E', 'Resultado'")
    Conn.Execute strSQL
    
    ''2.3.2.02 Reservas de Lucros C
    'strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsSaldos!APURACAO, ",", ".") & " AS VALOR, '" & "2.3.2.02" & "' AS CONTA, 'C' AS TIPO, 'Lucros Acumulados no Período - Resultado' AS HISTORICO, 'AUT', 'E', 'Resultado'")
    'Conn.Execute strSQL
    
    
    'Encerramento de Contas de Depesas como C
    Set rsResultado = Db.OpenRecordset("SELECT EFD_I155_Detalhe_Saldos.CHAVE, EFD_I155_Detalhe_Saldos.ANOMES, EFD_I155_Detalhe_Saldos.Cod_Conta, EFD_I050_Contas_Contabeis.DESCRICAO, EFD_I050_Contas_Contabeis.NATUREZA, EFD_I155_Detalhe_Saldos.Saldo_Final, EFD_I155_Detalhe_Saldos.Ind_Saldo_Fim " & _
    "FROM EFD_I050_Contas_Contabeis INNER JOIN EFD_I155_Detalhe_Saldos ON EFD_I050_Contas_Contabeis.COD = EFD_I155_Detalhe_Saldos.Cod_Conta " & _
    "WHERE (((EFD_I155_Detalhe_Saldos.ANOMES)=" & cAnomes & ") AND (Left(Cod_Conta,1)='3') AND ((EFD_I050_Contas_Contabeis.NATUREZA)='04'));")
    Do Until rsResultado.EOF
    'Manda pra Contas antes do resultado
    'strSQL = ("DELETE FROM EFD_I350_Detalhe_Saldos_Antes_Encerramento where CHAVE = '" & rsResultado!CHAVE & "'")
    
    strSQL = ("DELETE FROM EFD_I350_Detalhe_Saldos_Antes_Encerramento where CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    
    
    strSQL = ("INSERT INTO EFD_I350_Detalhe_Saldos_Antes_Encerramento (CHAVE,ANOMES,Data_INI,Data_FIM,Cod_Conta,Ind_Saldo_Ini,Saldo_Inicial,Total_Debitos,Total_Creditos,Saldo_Final,Ind_Saldo_Fim,Diferenca) SELECT * FROM EFD_I155_Detalhe_Saldos WHERE CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    
    'Faz o Lançamento Zerando
       
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsResultado!Saldo_Final * -1, ",", ".") & " AS VALOR, '" & rsResultado!Cod_Conta & "' AS CONTA, 'C' AS TIPO, 'Encerramento de contas de Receitas' AS HISTORICO, 'AUT', 'E', 'Resultado'")
    Conn.Execute strSQL
    
    cValorC = cValorC + (rsResultado!Saldo_Final * -1)
    rsResultado.MoveNext
    Loop
        
    'Encerramento de Contas de Receitas como D
    Set rsResultado = Db.OpenRecordset("SELECT EFD_I155_Detalhe_Saldos.CHAVE, EFD_I155_Detalhe_Saldos.ANOMES, EFD_I155_Detalhe_Saldos.Cod_Conta, EFD_I050_Contas_Contabeis.DESCRICAO, EFD_I050_Contas_Contabeis.NATUREZA, EFD_I155_Detalhe_Saldos.Saldo_Final, EFD_I155_Detalhe_Saldos.Ind_Saldo_Fim " & _
    "FROM EFD_I050_Contas_Contabeis INNER JOIN EFD_I155_Detalhe_Saldos ON EFD_I050_Contas_Contabeis.COD = EFD_I155_Detalhe_Saldos.Cod_Conta " & _
    "WHERE (((EFD_I155_Detalhe_Saldos.ANOMES)=" & cAnomes & ") AND (Left(Cod_Conta,1)='4') AND ((EFD_I050_Contas_Contabeis.NATUREZA)='04'));")
    Do Until rsResultado.EOF
    'Manda pra Contas antes do resultado
    
    
    strSQL = ("DELETE FROM EFD_I350_Detalhe_Saldos_Antes_Encerramento where CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    
    strSQL = ("INSERT INTO EFD_I350_Detalhe_Saldos_Antes_Encerramento (CHAVE,ANOMES,Data_INI,Data_FIM,Cod_Conta,Ind_Saldo_Ini,Saldo_Inicial,Total_Debitos,Total_Creditos,Saldo_Final,Ind_Saldo_Fim,Diferenca) SELECT * FROM EFD_I155_Detalhe_Saldos WHERE   CHAVE = '" & rsResultado!CHAVE & "'")
    Conn.Execute strSQL
    
    'Faz o Lançamento Zerando
    strSQL = ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", '" & Format(cDtFim, "yyyy-mm-dd") & "' AS DATA, " & Replace(rsResultado!Saldo_Final, ",", ".") & " AS VALOR, '" & rsResultado!Cod_Conta & "' AS CONTA, 'D' AS TIPO, 'Encerramento de contas de Despesas' AS HISTORICO, 'AUT', 'E', 'Resultado'")
    Conn.Execute strSQL
    
    cValorD = cValorD + (rsResultado!Saldo_Final)
    rsResultado.MoveNext
    Loop
    'atualiza o head
    strSQL = ("UPDATE EFD_I200_Lancamento_Contabil_Head SET Valor = " & Replace(Round(cValorD, 2), ",", ".") & " where EFD_I200_Lancamento_Contabil_Head.ID = " & cId & " ")
    Conn.Execute strSQL
    cValorC = 0
    cValorD = 0
    cValorT = 0
    'atualiza o head
    
    Else
    End If
    
    

rsSaldos.MoveNext
iCounter = iCounter + 1
      If iCounter = 100 Then
      DoEvents
      iCounter = 0
      End If
Loop

'ADICIONA ANOMES NOS LANÇAMENTOS
strSQL = ("UPDATE EFD_I200_Lancamento_Contabil " & _
"Set EFD_I200_Lancamento_Contabil.AnoMes = concat(Year(EFD_I200_Lancamento_Contabil.Data), Lpad(Month(EFD_I200_Lancamento_Contabil.Data), 2, 0)) " & _
"WHERE EFD_I200_Lancamento_Contabil.AnoMes is null;")
Call ConnectToDataBase
Conn.Execute strSQL

Call DisconnectFromDataBase


End Sub

Private Sub Resultado_Exercicio_Anterior(cAnomes)
Dim Db As Database
Set Db = CurrentDb()

'LANÇAMENTO RESULTADO EXERCÍCIOS ANTERIORES - LANÇAMENTO FUTURO PROXIMO PERIODO
'2.3.3.01 Lucros ou Prejuízos de Exercícios Anteriores
cAno = Left(cAnomes, 4)
cMes = Right(cAnomes, 2)
cMes = cMes + 1
If cMes > 12 Then
cMes = "01"
cAno = cAno + 1
Else
End If
cAnomesFuturo = cAno & Format(cMes, "00")

'Deleta caso já exista para evitar duplicidade
'DoCmd.RunSQL ("DELETE * FROM EFD_I200_Lancamento_Contabil where anomes = " & cAnomesFuturo & " and Conta = '2.3.3.01'")
Set rsSaldos = Db.OpenRecordset("select * from EFD_I155_Detalhe_Saldos where Cod_Conta = '2.3.3.02' and ANOMES = " & cAnomes & "")
   
'consulta data INI AnoMes futuro
Set rsCalendario = Db.OpenRecordset("select * from EFD_I150_Calendario where anomes = " & cAnomesFuturo & "")
   'HEAD
 '   DoCmd.RunSQL ("INSERT INTO EFD_I200_Lancamento_Contabil_Head (Data,Valor,Operacao,Indicador) SELECT #" & format(rsCalendario!DT_INI, "mm/dd/yyyy") & "# AS DATA, " & Replace(rsSaldos!Saldo_Final, ",", ".") & " AS VALOR, 'ENC', 'N'")
  '  Set rsHead = db.OpenRecordset("SELECT EFD_I200_Lancamento_Contabil_Head.ID FROM EFD_I200_Lancamento_Contabil_Head;")
  '  rsHead.MoveLast
  '  cId = rsHead!ID
    'HEAD
    '2.3.3.01 Lucros ou Prejuízos de Exercícios Anteriores
  '  If rsSaldos!Saldo_Final >= 0 Then
  '  DoCmd.RunSQL ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", #" & format(rsCalendario!DT_INI, "mm/dd/yyyy") & "# AS DATA, " & Replace(rsSaldos!Saldo_Final, ",", ".") & " AS VALOR, '" & "2.3.3.01" & "' AS CONTA, 'C' AS TIPO, 'Lucros Acumulados no Período Anterior - Resultado' AS HISTORICO, 'ENC', 'N', 'Resultado'")
  '  Else
  '  DoCmd.RunSQL ("INSERT INTO EFD_I200_Lancamento_Contabil (Id, Data,Valor,Conta, Tipo,Historico, OPERACAO, Indicador, Num_Nota) SELECT " & cId & ", #" & format(rsCalendario!DT_INI, "mm/dd/yyyy") & "# AS DATA, " & Replace(rsSaldos!Saldo_Final * -1, ",", ".") & " AS VALOR, '" & "2.3.3.01" & "' AS CONTA, 'D' AS TIPO, 'Prejuízos Acumulados no Período Anterior - Resultado' AS HISTORICO, 'ENC', 'N', 'Resultado'")
  '  End If


End Sub

Private Sub Refresh_Saldos_Fechamento(cAnomes)

Call ConnectToDataBase

Dim Db As Database
Set Db = CurrentDb()

Set rsContas = Db.OpenRecordset("SELECT COD_CONTA FROM EFD_I155_Detalhe_Saldos where anomes = " & cAnomes & " and (left(Cod_Conta,1)=3 or left(Cod_Conta,1)=4 or Cod_Conta = '2.3.3.02' or Cod_Conta = '2.3.3.03' or Cod_Conta = '2.3.2.02') group by COD_Conta")

strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Saldo_Inicial = 0, Saldo_Final = 0, Ind_Saldo_INI = 0, Ind_Saldo_FIM = 0, Diferenca = 0 where anomes = " & cAnomes & " and (left(Cod_Conta,1)=3 or left(Cod_Conta,1)=4 or Cod_Conta = '2.3.3.02' or Cod_Conta = '2.3.3.03' or Cod_Conta = '2.3.2.02')")
Conn.Execute strSQL

cAno = Left(cAnomes, 4)
cMes = Right(cAnomes, 2)
cMes = Int(cMes - 1)
If cMes = 0 Then
cMes = 12
cAno = cAno - 1
Else
End If
cAnomesAnt = cAno & Format(cMes, "00")

    Do Until rsContas.EOF = True
    cSaldoIni = 0
    cSaldoFim = 0
    
    Set rsContasHist = Db.OpenRecordset("SELECT * FROM EFD_I155_Detalhe_Saldos WHERE COD_CONTA = '" & rsContas!Cod_Conta & "' and anomes= " & cAnomes & "")
    Do Until rsContasHist.EOF
    
    cSaldoIni = cSaldoIni + rsContasHist!Saldo_Inicial
    cSaldoFim = cSaldoIni + rsContasHist!Total_Creditos - rsContasHist!Total_Debitos
    
    strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Saldo_Inicial = " & Replace(cSaldoIni, ",", ".") & ", Saldo_Final = " & Replace(cSaldoFim, ",", ".") & ", Diferenca = " & Replace(cSaldoIni - cSaldoFim, ",", ".") & " where CHAVE = '" & rsContasHist!ANOMES & rsContasHist!Cod_Conta & "'")
    Conn.Execute strSQL
   ' DoCmd.RunSQL ("UPDATE EFD_I155_Detalhe_Saldos_Antes_Encerramento SET Saldo_Inicial = " & Replace(cSaldoIni, ",", ".") & ", Saldo_Final = " & Replace(cSaldoFim, ",", ".") & ", Diferenca = " & Replace(cSaldoIni - cSaldoFim, ",", ".") & " where CHAVE = '" & rsContasHist!ANOMES & rsContasHist!Cod_Conta & "'")
    
    cSaldoIni = cSaldoFim
    rsContasHist.MoveNext
    Loop
    rsContas.MoveNext
    iCounter = iCounter + 1
      If iCounter = 20 Then
      DoEvents
      iCounter = 0
      End If
Loop

strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Ini = 'C' where saldo_Inicial >= 0 and anomes>= " & cAnomesAnt & " and anomes <= " & cAnomes & "")
Conn.Execute strSQL
strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Ini = 'D' where saldo_Inicial < 0 and anomes>= " & cAnomesAnt & " and anomes <= " & cAnomes & "")
Conn.Execute strSQL
strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Fim = 'C' where saldo_Final >= 0 and anomes>= " & cAnomesAnt & " and anomes <= " & cAnomes & "")
Conn.Execute strSQL
strSQL = ("UPDATE EFD_I155_Detalhe_Saldos SET Ind_Saldo_Fim = 'D' where saldo_Final < 0 and anomes>= " & cAnomesAnt & " and anomes <= " & cAnomes & "")
Conn.Execute strSQL

Call DisconnectFromDataBase

End Sub


Public Sub Contabilizacao_Saldos_Periodicos_Fechamento(cAnomes)

'DoCmd.setwarnings (False)

Dim Db As Database
Set Db = CurrentDb()
Call ConnectToDataBase



Dim cSaldoIni As Double
Dim cSaldoFim As Double

cSaldoIni = 0
cSaldoFim = 0


  Set rsI150 = Db.OpenRecordset("SELECT * FROM EFD_I150_Calendario where anomes =" & cAnomes & "")
   Dim cSaldoD As Double
   Dim cSaldoC As Double
   
  Do Until rsI150.EOF
  'DoCmd.RunSQL ("delete * from EFD_I155_Detalhe_Saldos where anomes = " & rsI150!Anomes & " and (left(Cod_Conta,1) = 3 or left(Cod_Conta,1) = 4 or Cod_conta = '2.3.3.03' or Cod_conta = '2.3.3.02' or Cod_conta = '2.3.2.02')")
  'o delete tem que ser especifico da chave porque caso nao tenha movimento perde o registro do ano passo
    
  'DEBITOS E CREDITOS
  Set rsSaldos = Db.OpenRecordset("TRANSFORM Sum(EFD_I200_Lancamento_Contabil.Valor) as Valor " & _
    "SELECT EFD_I200_Lancamento_Contabil.Conta " & _
    "FROM EFD_I200_Lancamento_Contabil " & _
    "WHERE (((Left(conta,1))=3 Or (Left(conta,1))=4) AND ((EFD_I200_Lancamento_Contabil.Data)>=#" & Format(rsI150!DT_INI, "mm/dd/yyyy") & "# And (EFD_I200_Lancamento_Contabil.Data)<=#" & Format(rsI150!DT_FIM, "mm/dd/yyyy") & "#)) OR (((EFD_I200_Lancamento_Contabil.Data)>=#" & Format(rsI150!DT_INI, "mm/dd/yyyy") & "# And (EFD_I200_Lancamento_Contabil.Data)<=#" & Format(rsI150!DT_FIM, "mm/dd/yyyy") & "#) AND ((EFD_I200_Lancamento_Contabil.conta)='2.3.3.03' Or (EFD_I200_Lancamento_Contabil.conta)='2.3.3.02' Or (EFD_I200_Lancamento_Contabil.conta)='2.3.2.02')) " & _
    "GROUP BY EFD_I200_Lancamento_Contabil.Conta " & _
    "PIVOT EFD_I200_Lancamento_Contabil.Tipo;")
    '"where EFD_I200_Lancamento_Contabil.Data >= #" & format(rsI150!DT_INI, "mm/dd/yyyy") & "# and EFD_I200_Lancamento_Contabil.Data <= #" & format(rsI150!DT_FIM, "mm/dd/yyyy") & "# and (left(conta,1)=3 or left(conta,1) = 4 or conta = '2.3.3.03' or conta = '2.3.3.02' or conta = '2.3.2.02') " & _

    Do Until rsSaldos.EOF
    cSaldoD = 0
    cSaldoC = 0
    'On Error Resume Next
    If IsNull(rsSaldos.Fields("D").Value) Then
    Else
    cSaldoD = rsSaldos!D
    End If
    If IsNull(rsSaldos.Fields("C").Value) Then
    Else
    cSaldoC = rsSaldos!c
    End If
    'faz o delete aqui
    
    strSQL = ("delete from EFD_I155_Detalhe_Saldos where chave = '" & rsI150!ANOMES & rsSaldos!Conta & "'")
    Conn.Execute strSQL
    
    strSQL = ("INSERT INTO EFD_I155_Detalhe_Saldos (CHAVE, ANOMES, Data_INI, Data_FIM, Cod_Conta, Total_Debitos, Total_Creditos) " & _
    "SELECT '" & rsI150!ANOMES & rsSaldos!Conta & "' as CHAVE, '" & rsI150!ANOMES & "' as ANOMES, " & _
    "'" & Format(rsI150!DT_INI, "yyyy-mm-dd") & "' as DTINI, " & _
    "'" & Format(rsI150!DT_FIM, "yyyy-mm-dd") & "' as DTFIM, " & _
    "'" & rsSaldos!Conta & "' as CONTA, " & _
    "" & Replace(cSaldoD, ",", ".") & " as D, " & _
    "" & Replace(cSaldoC, ",", ".") & " as C;")
    Conn.Execute strSQL
    
    rsSaldos.MoveNext
    iCounter = iCounter + 1
      If iCounter = 100 Then
      DoEvents
      iCounter = 0
      End If
    Loop
    
     'DEBITOS E CREDITOS DE CONTAS COM SALDO DO PERÍODO ANTERIOR SEM MOVIMENTO NESSE PERIODO
    Set rsSaldos = Db.OpenRecordset("TRANSFORM Sum(EFD_I200_Lancamento_Contabil.Valor) AS Valor " & _
        "SELECT EFD_I200_Lancamento_Contabil.Conta " & _
        "FROM EFD_I200_Lancamento_Contabil " & _
        "WHERE (((Left(conta,1))=3 Or (Left(conta,1))=4) AND ((EFD_I200_Lancamento_Contabil.Data)>=#" & Format(rsI150!DT_INI, "mm/dd/yyyy") & "# And (EFD_I200_Lancamento_Contabil.Data)<=#" & Format(rsI150!DT_FIM, "mm/dd/yyyy") & "#)) OR (((EFD_I200_Lancamento_Contabil.Data)>=#" & Format(rsI150!DT_INI, "mm/dd/yyyy") & "# And (EFD_I200_Lancamento_Contabil.Data)<=#" & Format(rsI150!DT_FIM, "mm/dd/yyyy") & "#) AND ((EFD_I200_Lancamento_Contabil.conta)='2.3.3.03' Or (EFD_I200_Lancamento_Contabil.conta)='2.3.3.02' Or (EFD_I200_Lancamento_Contabil.conta)='2.3.2.02')) " & _
        "GROUP BY EFD_I200_Lancamento_Contabil.Conta " & _
        "PIVOT EFD_I200_Lancamento_Contabil.Tipo;")
      '"WHERE (((EFD_I200_Lancamento_Contabil.Data) < #" & format(rsI150!DT_INI, "mm/dd/yyyy") & "# and (left(conta,1)=3 or left(conta,1) = 4 or conta = '2.3.3.03' or conta = '2.3.3.02' or conta = '2.3.2.02')) " & _

    Do Until rsSaldos.EOF
    cSaldoD = 0
    cSaldoC = 0
    If IsNull(rsSaldos.Fields("D").Value) Then
    Else
    cSaldoD = rsSaldos!D
    End If
    If IsNull(rsSaldos.Fields("C").Value) Then
    Else
    cSaldoC = rsSaldos!c
    End If
    
    cSaldoIni = cSaldoD + cSaldoC
    cSaldoFim = cSaldoIni
    
    Set rsTest = Db.OpenRecordset("select CHAVE from EFD_I155_Detalhe_Saldos where CHAVE = '" & rsI150!ANOMES & rsSaldos!Conta & "'")
    If rsTest.RecordCount = 0 Then
    
    'faz o delete aqui
    strSQL = ("delete * from EFD_I155_Detalhe_Saldos where chave = " & rsI150!ANOMES & rsSaldos!Conta & "")
    Conn.Execute strSQL
    
    strSQL = ("INSERT INTO EFD_I155_Detalhe_Saldos (CHAVE, ANOMES, Data_INI, Data_FIM, Cod_Conta, Total_Debitos, Total_Creditos) " & _
    "SELECT '" & rsI150!ANOMES & rsSaldos!Conta & "' as CHAVE, '" & rsI150!ANOMES & "' as ANOMES, " & _
    "'" & Format(rsI150!DT_INI, "yyyy-mm-dd") & "' as DTINI, " & _
    "'" & Format(rsI150!DT_FIM, "yyyy-mm-dd") & "' as DTFIM, " & _
    "'" & rsSaldos!Conta & "' as CONTA, " & _
    "" & Replace(cSaldoIni, ",", ".") & " as D, " & _
    "" & Replace(cSaldoFim, ",", ".") & " as C;")
    Conn.Execute strSQL
    
    Else
    End If
    
    rsSaldos.MoveNext
    iCounter = iCounter + 1
      If iCounter = 100 Then
      DoEvents
      iCounter = 0
      End If
    Loop
    
    
    rsI150.MoveNext
    
Loop
Call DisconnectFromDataBase

    
    End Sub



