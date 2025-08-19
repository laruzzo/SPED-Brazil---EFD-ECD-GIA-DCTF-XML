Attribute VB_Name = "modGeraArquivo_DCTF"
Option Compare Database

Public Function Gerar_DCTF(cDtIni As String, cDtFim As String, clocal As String, cDtINI_Contabil As String)

'ARQUIVO DCTF
'DoCmd.setwarnings (False)

'EXPORTAR ARQUIVO TXT
Dim iArq As Long
iArq = FreeFile

Open clocal & "\DCTF_" & month(cDtIni) & "_" & year(cDtIni) & ".dec" For Output As iArq

Dim cAno As String
Dim cMes As String
cAno = year(cDtIni)
cMes = month(cDtIni)
'23866944000141-DCTFM34-201810-ORIGI.dec

'Print #iArq, c0000 & Chr(13); c0001 & Chr(13); c0005 & Chr(13); c0100 & Chr(13); c0150 & Chr(13) & c0190 & Chr(13) & c0200 & Chr(13) & c0300 & Chr(13) & c0305 & Chr(13) & c0400 & Chr(13) & c0500 & Chr(13); c0600 & Chr(13) & c0990 & Chr(13) & cC001 & Chr(13) & cC100 & Chr(13) & cC170 & Chr(13) & cC190 & Chr(13) & cC500 & Chr(13) & cC501 & Chr(13) & cC990 & Chr(13) & cD001 & Chr(13) & cD190 & Chr(13) & cD990 & Chr(13) & cE001 & Chr(13) & cE100
'Print #iArq, c0000
Dim cSTR_DtINI As String
Dim cSTR_DtFIM As String

cSTR_DtINI = Replace(Format(cDtIni, "dd/mm/yyyy"), "/", "")
cSTR_DtFIM = Replace(Format(cDtFim, "dd/mm/yyyy"), "/", "")


cDtIni = Format(cDtIni, "mm/dd/yyyy")
cDtFim = Format(cDtFim, "mm/dd/yyyy")

cAno = Right(Left(cSTR_DtINI, 8), 4)
cMes = Int(Right(Left(cSTR_DtINI, 4), 2))

Dim Db As Database
Set Db = CurrentDb()

Dim rsEmpresa As DAO.Recordset
Set rsEmpresa = Db.OpenRecordset("tbEmpresa")

Dim rsContador As DAO.Recordset
Set rsContador = Db.OpenRecordset("tbContador")

Dim rsCalendario As DAO.Recordset
Set rsCalendario = Db.OpenRecordset("select * from tbResumo_IRPJ_CSLL where ano = '" & cAno & "' and mes = " & cMes & "")

Dim rsCliente As DAO.Recordset
Set rsCliente = Db.OpenRecordset("SELECT tbCliente.*, tbVendas.DataEmissao " & _
"FROM tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente " & _
"WHERE (((tbVendas.DataEmissao)>=# " & cDtIni & "  # And (tbVendas.DataEmissao)<=# " & cDtFim & " #));")

Dim rsIPI As DAO.Recordset
Set rsIPI = Db.OpenRecordset("SELECT * FROM tbResumo_IPI " & _
"WHERE ANO = '" & cAno & "'  And MES = " & cMes & "")


Dim rsPIS As DAO.Recordset
Set rsPIS = Db.OpenRecordset("SELECT * FROM tbResumo_PIS " & _
"WHERE ANO = '" & cAno & "'  And MES = " & cMes & "")

Dim rsCofins As DAO.Recordset
Set rsCofins = Db.OpenRecordset("SELECT * FROM tbResumo_Cofins " & _
"WHERE ANO = '" & cAno & "'  And MES = " & cMes & "")



Dim cCount As Integer
cCount = 0

'HEADER
Dim h1 As String * 5
Dim h2 As String * 3
Dim h3 As String * 4
Dim h4 As String * 4
Dim h5 As String * 4
Dim h6 As String * 1
Dim h7 As String * 14
Dim h8 As String * 1
Dim h9 As String * 3
Dim h10 As String * 60
Dim h11 As String * 2
Dim h12 As String * 10
Dim h13 As String * 1
Dim h14 As String * 2
Dim h15 As String * 4
Dim h16 As String * 2
Dim h17 As String * 11
Dim h18 As String * 8
Dim h19 As String * 8
Dim h20 As String * 8
Dim h21 As String * 1
Dim h22 As String * 1
Dim h23 As String * 207
Dim h24 As String * 10
'Dim h25 As String * 2

h1 = "DCTFM"
h2 = ""
h3 = ""
h4 = Right(Left(cSTR_DtINI, 8), 4)
h5 = "1930"
h6 = "0"
h7 = rsEmpresa!CNPJ
h8 = "0"
h9 = "340"
h10 = rsEmpresa!RazaoSocial
h11 = rsEmpresa!UF
h12 = "0000000000"
h13 = "0"
h14 = "00"
h15 = Right(Left(cSTR_DtINI, 8), 4)
h16 = Right(Left(cSTR_DtINI, 4), 2)
h17 = "00000000000"
h18 = Left(cSTR_DtINI, 8)
h19 = Left(cSTR_DtFIM, 8)
h20 = "00000000"
h21 = "0"
h22 = "0"
h23 = ""
h24 = "0000000000"
'h25 = "0D"

cHeader = h1 & h2 & h3 & h4 & h5 & h6 & h7 & h8 & h9 & h10 & h11 & h12 & h13 & h14 & h15 & h16 & h17 & h18 & h19 & h20 & h21 & h22 & h23 & h24 '& h25
Print #iArq, cHeader
cCount = cCount + 1

'HEADER

'Dados Iniciais - Tipo R01
Dim r1_1 As String * 3
Dim r1_2 As String * 14
Dim r1_3 As String * 6
Dim r1_4 As String * 1
Dim r1_5 As String * 8
Dim r1_6 As String * 4
Dim r1_7 As String * 4
Dim r1_8 As String * 1
Dim r1_9 As String * 12
Dim r1_10 As String * 1
Dim r1_11 As String * 2
Dim r1_12 As String * 1
Dim r1_13 As String * 1
Dim r1_14 As String * 1
Dim r1_15 As String * 1
Dim r1_16 As String * 1
Dim r1_17 As String * 1
Dim r1_18 As String * 11
Dim r1_19 As String * 1
Dim r1_20 As String * 1
Dim r1_21 As String * 1
Dim r1_22 As String * 10
'Dim r1_23 As String * 2


r1_1 = "R01"
r1_2 = Left(rsEmpresa!CNPJ, 14)
r1_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r1_4 = "0"
r1_5 = "00000000"
r1_6 = Left(cSTR_DtINI, 2) & Right(Left(cSTR_DtINI, 4), 2)
r1_7 = Left(cSTR_DtFIM, 2) & Right(Left(cSTR_DtFIM, 4), 2)
r1_8 = "0"
r1_9 = "000000000000"
r1_10 = "0"
r1_11 = "07"
r1_12 = "0"
r1_13 = "0"
r1_14 = "0"
r1_15 = "0"
r1_16 = "0"
r1_17 = "4"
r1_18 = "00000000000"
r1_19 = "3"
r1_20 = "4"
r1_21 = "0"
r1_22 = ""
'r1_23 = "0D"

cR01 = r1_1 & r1_2 & r1_3 & r1_4 & r1_5 & r1_6 & r1_7 & r1_8 & r1_9 & r1_10 & r1_11 & r1_12 & r1_13 & r1_14 & r1_15 & r1_16 & r1_17 & r1_18 & r1_19 & r1_20 & r1_21 & r1_22 '& r1_23
Print #iArq, cR01
cCount = cCount + 1
'Dados Iniciais - Tipo R01

'Dados Cadastrais do Estabelecimento Matriz - Tipo R02
Dim r2_1 As String * 3
Dim r2_2 As String * 14
Dim r2_3 As String * 6
Dim r2_4 As String * 1
Dim r2_5 As String * 8
Dim r2_6 As String * 115
Dim r2_7 As String * 4
Dim r2_8 As String * 40
Dim r2_9 As String * 6
Dim r2_10 As String * 21
Dim r2_11 As String * 20
Dim r2_12 As String * 50
Dim r2_13 As String * 2
Dim r2_14 As String * 8
Dim r2_15 As String * 4
Dim r2_16 As String * 9
Dim r2_17 As String * 4
Dim r2_18 As String * 9
Dim r2_19 As String * 6
Dim r2_20 As String * 2
Dim r2_21 As String * 8
Dim r2_22 As String * 40
Dim r2_23 As String * 10
'Dim r2_24 As String * 2

r2_1 = "R02"
r2_2 = rsEmpresa!CNPJ
r2_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r2_4 = "0"
r2_5 = "00000000"
r2_6 = rsEmpresa!RazaoSocial
r2_7 = "0000"
r2_8 = rsEmpresa!Logradouro
r2_9 = rsEmpresa!Num
r2_10 = rsEmpresa!compl
r2_11 = rsEmpresa!Bairro
r2_12 = rsEmpresa!Municipio
r2_13 = rsEmpresa!UF
r2_14 = rsEmpresa!CEP
r2_15 = "00" + Left(rsEmpresa!Fone, 2)
r2_16 = Right(Left(rsEmpresa!Fone, 11), 9)
r2_17 = ""
r2_18 = ""
r2_19 = ""
r2_20 = ""
r2_21 = ""
r2_22 = "cezar.barreto@overturecervejaria.com.br"
r2_23 = ""
'r2_24 = "0D"

CR02 = r2_1 & r2_2 & r2_3 & r2_4 & r2_5 & r2_6 & r2_7 & r2_8 & r2_9 & r2_10 & r2_11 & r2_12 & r2_13 & r2_14 & r2_15 & r2_16 & r2_17 & r2_18 & r2_19 & r2_20 & r2_21 & r2_22 & r2_23 '& r2_24
Print #iArq, CR02
cCount = cCount + 1
'Dados Cadastrais do Estabelecimento Matriz - Tipo R02

'Dados dos Responsáveis pela Pessoa Jurídica - Tipo R03
Dim r3_1 As String * 3
Dim r3_2 As String * 14
Dim r3_3 As String * 6
Dim r3_4 As String * 1
Dim r3_5 As String * 8
Dim r3_6 As String * 60
Dim r3_7 As String * 11
Dim r3_8 As String * 4
Dim r3_9 As String * 9
Dim r3_10 As String * 5
Dim r3_11 As String * 4
Dim r3_12 As String * 9
Dim r3_13 As String * 40
Dim r3_14 As String * 60
Dim r3_15 As String * 11
Dim r3_16 As String * 15
Dim r3_17 As String * 2
Dim r3_18 As String * 4
Dim r3_19 As String * 9
Dim r3_20 As String * 5
Dim r3_21 As String * 4
Dim r3_22 As String * 9
Dim r3_23 As String * 40
Dim r3_24 As String * 10
'Dim r3_25 As String * 2

r3_1 = "R03"
r3_2 = rsEmpresa!CNPJ
r3_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r3_4 = "0"
r3_5 = "00000000"
r3_6 = rsContador!NomeContador
r3_7 = rsContador!CPFContador
r3_8 = Left(rsContador!TelefoneEscritorio, 2)
r3_9 = Right(rsContador!TelefoneEscritorio, 9)
r3_10 = ""
r3_11 = ""
r3_12 = ""
r3_13 = rsContador!EmailEscritorio
r3_14 = rsContador!NomeContador
r3_15 = rsContador!CPFContador
r3_16 = ""
r3_17 = ""
r3_18 = Left(rsContador!TelefoneEscritorio, 2)
r3_19 = Right(rsContador!TelefoneEscritorio, 9)
r3_20 = ""
r3_21 = ""
r3_22 = ""
r3_23 = rsContador!EmailEscritorio
r3_24 = ""
'r3_25 = "0D"

cR03 = r3_1 & r3_2 & r3_3 & r3_4 & r3_5 & r3_6 & r3_7 & r3_8 & r3_9 & r3_10 & r3_11 & r3_12 & r3_13 & r3_14 & r3_15 & r3_16 & r3_17 & r3_18 & r3_19 & r3_20 & r3_21 & r3_22 & r3_23 & r3_24 '& r3_25
Print #iArq, cR03
cCount = cCount + 1

'Dados dos Responsáveis pela Pessoa Jurídica - Tipo R03

'Débito Apurado e Créditos Vinculados - Tipo R10
Dim r10_1 As String * 3
Dim r10_2 As String * 14
Dim r10_3 As String * 6
Dim r10_4 As String * 1
Dim r10_5 As String * 8
Dim r10_6 As String * 2
Dim r10_7 As String * 6
Dim r10_8 As String * 1
Dim r10_9 As String * 4
Dim r10_10 As String * 2
Dim r10_11 As String * 2
Dim r10_12 As String * 6
Dim r10_13 As String * 14
Dim r10_14 As String * 1
Dim r10_15 As String * 14
Dim r10_16 As String * 1
Dim r10_17 As String * 1
Dim r10_18 As String * 1
Dim r10_19 As String * 1
Dim r10_20 As String * 10
Dim r10_21 As String * 2

'IPI
If rsIPI!SALDO = 0 Then
GoTo semIPI
Else: End If
r10_1 = "R10"
r10_2 = rsEmpresa!CNPJ
r10_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r10_4 = "0"
'r10_5 = Replace(Format(cDtFIM, "dd/mm/yyyy"), "/", "")
r10_5 = "00000000"
r10_6 = "03" 'IPI
r10_7 = "066803" 'IPI-Bebidas do capitulo 22 da TIPI
r10_8 = "M"
r10_9 = Right(Left(cSTR_DtFIM, 8), 4)
r10_10 = Right(Left(cSTR_DtFIM, 4), 2)
r10_11 = "00"
r10_12 = "000141"
r10_13 = "00000000000000"
r10_14 = "0"
'r10_15 = rsIPI!Saldo
r10_15 = Right("00000000000000" & Replace(Round(rsIPI!SALDO * -1, 2), ",", ""), 14)
r10_16 = "0"
r10_17 = "0"
r10_18 = "0"
r10_19 = "0"
r10_20 = ""
'r10_21 = "0D"
If rsIPI!SALDO < 0 Then
cR10 = r10_1 & r10_2 & r10_3 & r10_4 & r10_5 & r10_6 & r10_7 & r10_8 & r10_9 & r10_10 & r10_11 & r10_12 & r10_13 & r10_14 & r10_15 & r10_16 & r10_17 & r10_18 & r10_19 & r10_20
Print #iArq, cR10
cCount = cCount + 1
Else
End If
semIPI:
'FIM IPI

'Débito Apurado e Créditos Vinculados - Tipo R10
'PIS
If rsPIS!DEB = 0 Then
GoTo semPIS
Else: End If
r10_1 = "R10"
r10_2 = rsEmpresa!CNPJ
r10_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r10_4 = "0"
'r10_5 = Replace(Format(cDtFIM, "dd/mm/yyyy"), "/", "")
r10_5 = "00000000"
r10_6 = "06" 'PIS PASEP
r10_7 = "067903" 'PIS - Tributação Bebidas Frias - Cervejas
r10_8 = "M"
r10_9 = Right(Left(cSTR_DtFIM, 8), 4)
r10_10 = Right(Left(cSTR_DtFIM, 4), 2)
r10_11 = "00"
r10_12 = "000000"
r10_13 = "00000000000000"
r10_14 = "0"
'r10_15 = Space(14 - Len(Round(rsPIS!DEB, 2))) & Round(rsPIS!DEB, 2)
'r10_15 = Pad(14 - Len(Round(rsPIS!DEB, 2))) & Round(rsPIS!DEB, 2)
r10_15 = Right("00000000000000" & Replace(Round(rsPIS!SALDO * -1, 2), ",", ""), 14)
r10_16 = "0"
r10_17 = "0"
r10_18 = "0"
r10_19 = "0"
r10_20 = ""
'r10_21 = "0D"

If rsPIS!SALDO < 0 Then
cR10 = r10_1 & r10_2 & r10_3 & r10_4 & r10_5 & r10_6 & r10_7 & r10_8 & r10_9 & r10_10 & r10_11 & r10_12 & r10_13 & r10_14 & r10_15 & r10_16 & r10_17 & r10_18 & r10_19 & r10_20
Print #iArq, cR10
cCount = cCount + 1
Else
End If
semPIS:
'FIM PIS

'Débito Apurado e Créditos Vinculados - Tipo R10
'COFINS
If rsCofins!DEB = 0 Then
GoTo semCofins
End If

r10_1 = "R10"
r10_2 = rsEmpresa!CNPJ
r10_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r10_4 = "0"
'r10_5 = Replace(Format(cDtFIM, "dd/mm/yyyy"), "/", "")
r10_5 = "00000000"
r10_6 = "07" 'COFINS
r10_7 = "076003" 'COFINS - Regime Especial de tributação - Cervejas
r10_8 = "M"
r10_9 = Right(Left(cSTR_DtFIM, 8), 4)
r10_10 = Right(Left(cSTR_DtFIM, 4), 2)
r10_11 = "00"
r10_12 = "000000"
r10_13 = "00000000000000"
r10_14 = "0"
'r10_15 = Space(14 - Len(Round(rsCofins!DEB, 2))) & Round(rsCofins!DEB, 2)
r10_15 = Right("00000000000000" & Replace(Round(rsCofins!SALDO * -1, 2), ",", ""), 14)
r10_16 = "0"
r10_17 = "0"
r10_18 = "0"
r10_19 = "0"
r10_20 = ""
'r10_21 = "0D"
If rsCofins!SALDO < 0 Then
cR10 = r10_1 & r10_2 & r10_3 & r10_4 & r10_5 & r10_6 & r10_7 & r10_8 & r10_9 & r10_10 & r10_11 & r10_12 & r10_13 & r10_14 & r10_15 & r10_16 & r10_17 & r10_18 & r10_19 & r10_20
Print #iArq, cR10
cCount = cCount + 1
Else
End If
semCofins:
'FIM COFINS

'VERIFICA SE O MÊS TEM IRPJ E CSLL
If rsCalendario.EOF = True And rsCalendario.BOF = True Then

'If IsNull(rsCalendario.Fields("Ano").Value) Then
Else
MsgBox ("Fechamento de trimestre detectado, tem que pagar IRPJ e CSLL. Verifique")
GoTo pulaIR:
'IRPJ
r10_1 = "R10"
r10_2 = rsEmpresa!CNPJ
r10_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r10_4 = "0"
'r10_5 = Replace(Format(cDtFIM, "dd/mm/yyyy"), "/", "")
r10_5 = "00000000"
r10_6 = "01" 'IRPJ
r10_7 = "208901" 'IRPJ - Lucro Presumido
r10_8 = "T"
r10_9 = Right(Left(cSTR_DtFIM, 8), 4) 'ano
Select Case cMes
Case Is = 3
r10_10 = "01"  'trimestre
Case Is = 6
r10_10 = "02"  'trimestre
Case Is = 9
r10_10 = "03"  'trimestre
Case Is = 12
r10_10 = "04"  'trimestre
End Select
r10_11 = "00"
r10_12 = "000000"
r10_13 = "00000000000000"
r10_14 = "0"
'r10_15 = Space(14 - Len(Round(rsCofins!DEB, 2))) & Round(rsCofins!DEB, 2)
r10_15 = Right("00000000000000" & Replace(Round(rsCalendario!VL_IRPJ, 2), ",", ""), 14)
r10_16 = "0"
r10_17 = "0"
r10_18 = "0"
r10_19 = "0"
r10_20 = ""
'r10_21 = "0D"
If rsCalendario!VL_IRPJ <= 0 Then
Else
cR10 = r10_1 & r10_2 & r10_3 & r10_4 & r10_5 & r10_6 & r10_7 & r10_8 & r10_9 & r10_10 & r10_11 & r10_12 & r10_13 & r10_14 & r10_15 & r10_16 & r10_17 & r10_18 & r10_19 & r10_20
Print #iArq, cR10
cCount = cCount + 1
End If

'CSLL
r10_1 = "R10"
r10_2 = rsEmpresa!CNPJ
r10_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
r10_4 = "0"
'r10_5 = Replace(Format(cDtFIM, "dd/mm/yyyy"), "/", "")
r10_5 = "00000000"
r10_6 = "05" 'CSLL
r10_7 = "237201" 'CSLL - Lucro Presumido ou Arbitrado
r10_8 = "T"
r10_9 = Right(Left(cSTR_DtFIM, 8), 4) 'ano
Select Case cMes
Case Is = 3
r10_10 = "01"  'trimestre
Case Is = 6
r10_10 = "02"  'trimestre
Case Is = 9
r10_10 = "03"  'trimestre
Case Is = 12
r10_10 = "04"  'trimestre
End Select
r10_11 = "00"
r10_12 = "000000"
r10_13 = "00000000000000"
r10_14 = "0"
'r10_15 = Space(14 - Len(Round(rsCofins!DEB, 2))) & Round(rsCofins!DEB, 2)
r10_15 = Right("00000000000000" & Replace(Round(rsCalendario!VL_CSLL, 2), ",", ""), 14)
r10_16 = "0"
r10_17 = "0"
r10_18 = "0"
r10_19 = "0"
r10_20 = ""
'r10_21 = "0D"
If rsCalendario!VL_CSLL <= 0 Then
Else
cR10 = r10_1 & r10_2 & r10_3 & r10_4 & r10_5 & r10_6 & r10_7 & r10_8 & r10_9 & r10_10 & r10_11 & r10_12 & r10_13 & r10_14 & r10_15 & r10_16 & r10_17 & r10_18 & r10_19 & r10_20
Print #iArq, cR10
cCount = cCount + 1
End If
End If


'VERIFICA SE O MÊS TEM IRPJ E CSLL
pulaIR:

'Trailler da Declaração - Tipo T9
Dim rT9_1 As String * 2
Dim rT9_2 As String * 14
Dim rT9_3 As String * 6
Dim rT9_4 As String * 1
Dim rT9_5 As String * 8
Dim rT9_6 As String * 5
Dim rT9_7 As String * 56
Dim rT9_8 As String * 10
'Dim rT9_9 As String * 2

rT9_1 = "T9"
rT9_2 = rsEmpresa!CNPJ
rT9_3 = Right(Left(cSTR_DtINI, 8), 4) & Right(Left(cSTR_DtINI, 4), 2)
rT9_4 = "0"
rT9_5 = Replace(Now, "/", "")
rT9_6 = cCount + 1
rT9_7 = ""
rT9_8 = ""
'rT9_9 = "0D"

cT9 = rT9_1 & rT9_2 & rT9_3 & rT9_4 & rT9_5 & rT9_6 & rT9_7 & rT9_8 '& rT9_9
Print #iArq, cT9;

'Trailler da Declaração - Tipo T9

Close #iArq
'DoCmd.setwarnings (True)
'MsgBox ("Arquivo Gerado")


End Function


