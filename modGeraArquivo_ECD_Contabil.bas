Attribute VB_Name = "modGeraArquivo_ECD_Contabil"
Option Compare Database


Public Conn As New ADODB.Connection
Public SQLStr As String

Public Sub ConnectToDataBase()
 
  Dim Server_Name As String
 Dim Database_Name As String
 Dim User_ID As String
 Dim Password As String
 
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

Public Function Gerar_ECD_Contabil(cDtIni As String, cDtFim As String, clocal As String, cDtINI_Contabil As String, cIDIventario As String)

'Registro 0000
'Registro I051
'Registro J150

'ARQUIVO ECD Contabil
'DoCmd.setwarnings (False)

Call ConnectToDataBase

Dim cLayout As String
'cLayout = "8.00"
'layout ref 2019 ECD 2020
cLayout = "9.00"


'EXPORTAR ARQUIVO TXT
Dim iArq As Long
Dim cPath As String
iArq = FreeFile

Open clocal & "\ECD_Contabil_" & month(cDtIni) & "_" & year(cDtIni) & ".txt" For Output As iArq
cPath = clocal & "\ECD_Contabil_" & month(cDtIni) & "_" & year(cDtIni) & ".txt"
'Print #iArq, c0000 & Chr(13); c0001 & Chr(13); c0005 & Chr(13); c0100 & Chr(13); c0150 & Chr(13) & c0190 & Chr(13) & c0200 & Chr(13) & c0300 & Chr(13) & c0305 & Chr(13) & c0400 & Chr(13) & c0500 & Chr(13); c0600 & Chr(13) & c0990 & Chr(13) & cC001 & Chr(13) & cC100 & Chr(13) & cC170 & Chr(13) & cC190 & Chr(13) & cC500 & Chr(13) & cC501 & Chr(13) & cC990 & Chr(13) & cD001 & Chr(13) & cD190 & Chr(13) & cD990 & Chr(13) & cE001 & Chr(13) & cE100
'Print #iArq, c0000
Dim cSTR_DtINI As String
Dim cSTR_DtFIM As String

cSTR_DtINI = Replace(Format(cDtIni, "dd/mm/yyyy"), "/", "")
cSTR_DtFIM = Replace(Format(cDtFim, "dd/mm/yyyy"), "/", "")

'NAO MUDE DE LUGAR O CALL
'Call PARTIDAS_DOBRADAS(cDtIni, cDtFim)
'Call Contabilizacao_Saldos_Periodicos(cDtIni, cDtFim)



cDtIni = Format(cDtIni, "yyyy-mm-dd")
cDtFim = Format(cDtFim, "yyyy-mm-dd")

cDtIniVb = Format(cDtIni, "mm/dd/yyyy")
cDtFimVb = Format(cDtFim, "mm/dd/yyyy")



Dim Db As Database
Set Db = CurrentDb()

Dim rsEmpresa As DAO.Recordset
Set rsEmpresa = Db.OpenRecordset("tbEmpresa")

Dim rsContador As DAO.Recordset
Set rsContador = Db.OpenRecordset("tbContador")

Dim rsCliente As DAO.Recordset
Set rsCliente = Db.OpenRecordset("SELECT tbCliente.*, tbVendas.DataEmissao " & _
"FROM tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente " & _
"WHERE (((tbVendas.DataEmissao)>=# " & cDtIniVb & "  # And (tbVendas.DataEmissao)<=# " & cDtFimVb & " #));")


'Limpa Fornecedores ativos
strSQL = ("delete from tbFornecedor_Ativo_temp")
Conn.Execute strSQL

'Insere Fornecedores de compras ativas
strSQL = ("insert into tbFornecedor_Ativo_temp " & _
"SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email " & _
"FROM tbFornecedor INNER JOIN (tbImobilizado INNER JOIN tbCompras ON tbImobilizado.ChaveNFe = tbCompras.ChaveNF) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
"WHERE (((tbImobilizado.DataEmissao) >= '" & cDtIni & "' And (tbImobilizado.DataEmissao) <= '" & cDtFim & "')) " & _
"GROUP BY tbCompras.IdFornecedor, tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email; ")
Conn.Execute strSQL

'Insere fornecedores de imobilizado ativos
strSQL = ("insert into tbFornecedor_Ativo_temp " & _
"SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email " & _
"FROM tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor LEFT OUTER JOIN tbFornecedor_Ativo_temp ON  tbFornecedor.IDFor = tbFornecedor_Ativo_temp.IDFor " & _
"WHERE tbCompras.DataEmissao >= '" & cDtIni & "' And tbCompras.DataEmissao <= '" & cDtFim & "' and tbFornecedor_Ativo_temp.IDFor is null " & _
"GROUP BY tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email;")
Conn.Execute strSQL


'insere transportadoras
strSQL = ("INSERT INTO tbFornecedor_Ativo_temp ( IDFor, Tipo, Cnpj, RazaoSocial, IE, CRT, CEP, Logradouro, Nro, Compl, Bairro, UF, cod_Municipio, Municipio, Pais, Fone, Email ) " & _
"SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email " & _
"FROM tbFornecedor INNER JOIN tbTransportes ON tbFornecedor.IDFor = tbTransportes.ID_Emit " & _
"GROUP BY tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email, tbTransportes.DataEmissao " & _
"HAVING (((tbTransportes.DataEmissao)>='" & cDtIni & "' And (tbTransportes.DataEmissao)<='" & cDtFim & "'));")
Conn.Execute strSQL


'Recordset Fornecedor Group By, remove dupLicados
Dim rsFornecedor As DAO.Recordset
Set rsFornecedor = Db.OpenRecordset("SELECT tbFornecedor_Ativo_temp.IDFor, tbFornecedor_Ativo_temp.Tipo, tbFornecedor_Ativo_temp.Cnpj, tbFornecedor_Ativo_temp.RazaoSocial, tbFornecedor_Ativo_temp.IE, tbFornecedor_Ativo_temp.CRT, tbFornecedor_Ativo_temp.CEP, tbFornecedor_Ativo_temp.Logradouro, tbFornecedor_Ativo_temp.Nro, tbFornecedor_Ativo_temp.Compl, tbFornecedor_Ativo_temp.Bairro, tbFornecedor_Ativo_temp.UF, tbFornecedor_Ativo_temp.cod_Municipio, tbFornecedor_Ativo_temp.Municipio, tbFornecedor_Ativo_temp.Pais, tbFornecedor_Ativo_temp.Fone, tbFornecedor_Ativo_temp.Email " & _
"FROM tbFornecedor_Ativo_temp " & _
"GROUP BY tbFornecedor_Ativo_temp.IDFor, tbFornecedor_Ativo_temp.Tipo, tbFornecedor_Ativo_temp.Cnpj, tbFornecedor_Ativo_temp.RazaoSocial, tbFornecedor_Ativo_temp.IE, tbFornecedor_Ativo_temp.CRT, tbFornecedor_Ativo_temp.CEP, tbFornecedor_Ativo_temp.Logradouro, tbFornecedor_Ativo_temp.Nro, tbFornecedor_Ativo_temp.Compl, tbFornecedor_Ativo_temp.Bairro, tbFornecedor_Ativo_temp.UF, tbFornecedor_Ativo_temp.cod_Municipio, tbFornecedor_Ativo_temp.Municipio, tbFornecedor_Ativo_temp.Pais, tbFornecedor_Ativo_temp.Fone, tbFornecedor_Ativo_temp.Email;")




Dim rsContas As DAO.Recordset
Set rsContas = Db.OpenRecordset("tbPlanoContasContabeis")



'TB TEMP CADASTRO PRODS ATIVOS
'limpa
strSQL = ("delete  from tbCadProd_Ativo_temp")
Conn.Execute strSQL

'vendas
strSQL = ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
"SELECT tbCadProd.IDProd " & _
"FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= '" & cDtIni & "' And (tbVendas.DataEmissao) <= '" & cDtFim & "')) " & _
"GROUP BY tbCadProd.IDProd;")
Conn.Execute strSQL

'delete de prods ativos os imobilizados com valor de icms = 0
'DoCmd.RunSQL ("DELETE DISTINCTROW tbCadProd_Ativo_temp.*, tbImobilizado.Valor_ICMS_total " & _
'"FROM tbImobilizado INNER JOIN tbCadProd_Ativo_temp ON tbImobilizado.IDProd = tbCadProd_Ativo_temp.IDProd " & _
'"WHERE (((tbImobilizado.Valor_ICMS_total)=0));")


'compras
strSQL = ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
"SELECT tbCadProd.IDProd " & _
"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= '" & cDtIni & "' And (tbCompras.DataEmissao) <= '" & cDtFim & "')) " & _
"GROUP BY tbCadProd.IDProd;")
Conn.Execute strSQL


Dim rsCFOP As DAO.Recordset
Set rsCFOP = Db.OpenRecordset("SELECT tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
"FROM tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbComprasDet INNER JOIN tbCadProd_Ativo_temp ON tbComprasDet.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
"GROUP BY tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbCompras.IdFornecedor " & _
"HAVING (((tbCompras.IdFornecedor)<>1131)); " & _
"UNION " & _
"SELECT tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC FROM tbVendasDet INNER JOIN tbCadProd_Ativo_temp ON tbVendasDet.IDProd = tbCadProd_Ativo_temp.IDProd " & _
"GROUP BY tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
"HAVING (((tbVendasDet.CFOP_ESCRITURADA) Is Not Null));")


'imobilizado
strSQL = ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
"SELECT tbCadProd.IDProd " & _
"FROM tbVendas, tbCadProd INNER JOIN tbImobilizado ON tbCadProd.IDProd = tbImobilizado.IDProd LEFT OUTER JOIN tbCadProd_Ativo_temp ON tbCadProd.IDProd = tbCadProd_Ativo_temp.IDProd  " & _
"WHERE tbImobilizado.DataEmissao >= '" & cDtIni & "' And tbImobilizado.DataEmissao <= '" & cDtFim & "' and tbCadProd_Ativo_temp.IDProd is null " & _
"GROUP BY tbCadProd.IDProd;")
Conn.Execute strSQL


'iventario
''DoCmd.setwarnings (True)
'DoCmd.RunSQL ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
'"SELECT tbIventarioDet.ID_Prod " & _
'"FROM tbIventarioDet LEFT JOIN tbCadProd_Ativo_temp ON tbIventarioDet.ID_Prod = tbCadProd_Ativo_temp.IDProd " & _
'"GROUP BY tbIventarioDet.ID_Prod, tbCadProd_Ativo_temp.IDProd, tbIventarioDet.ID_Iventario " & _
'"HAVING tbCadProd_Ativo_temp.IDProd AND tbIventarioDet.ID_Iventario=" & cIDIventario & ";")

strSQL = ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
"SELECT tbIventarioDet.ID_Prod " & _
"FROM tbIventarioDet LEFT JOIN tbCadProd_Ativo_temp ON tbIventarioDet.ID_Prod = tbCadProd_Ativo_temp.IDProd " & _
"GROUP BY tbCadProd_Ativo_temp.IDProd, tbIventarioDet.ID_Prod, tbIventarioDet.ID_Iventario " & _
"HAVING (((tbCadProd_Ativo_temp.IDProd) Is Null) AND ((tbIventarioDet.ID_Iventario)=" & cIDIventario & "));")
Conn.Execute strSQL



Dim rsMedidas As DAO.Recordset
Set rsMedidas = Db.OpenRecordset("SELECT tbCadProd.Unid, 'Med-' & Unid AS [Desc] FROM tbCadProd INNER JOIN tbCadProd_Ativo_temp ON tbCadProd.IDProd = tbCadProd_Ativo_temp.IDProd GROUP BY tbCadProd.Unid, 'Med-' & Unid;")


Dim rsCadProd As DAO.Recordset
Set rsCadProd = Db.OpenRecordset("SELECT tbCadProd.* FROM cstbCadProd_Ativo_temp INNER JOIN tbCadProd ON cstbCadProd_Ativo_temp.IDProd = tbCadProd.IDProd;")


Dim cCodVer As String
Dim cCodFin As String
Dim cPerfil As String
Dim cAtividade As String
Dim cPais As String
Dim cCNPJ As String
Dim cCPF As String
Dim cTipoItem As String
Dim cCusto As String
Dim cFunc As String
Dim clintot As Integer


'cSTR_DtINI = Replace(Format(cDtINI, "dd/mm/yyyy"), "/", "")
'cSTR_DtFIM = Replace(Format(cDtFIM, "dd/mm/yyyy"), "/", "")





cCodVer = "003"


cCodFin = "0"
cPerfil = "A"
cAtividade = "0"

clintot = 0


Dim c0000 As String
Dim c0001 As String
Dim c0005 As String
Dim c0015 As String
Dim c0100 As String
Dim c0150 As String
Dim c0190 As String
Dim c0200 As String
Dim c0205 As String
Dim c0206 As String
Dim c0210 As String
Dim c0220 As String
Dim c0300 As String
Dim c0305 As String
Dim c0400 As String
Dim c0450 As String
Dim c0500 As String

Dim clinB01 As Integer


'BLOCO 0: ABERTURA, IDENTIFICAÇÃO E REFERÊNCIAS.


'c0000
'REGISTRO 0:ABERTURA DO ARQUIVODIGITAL E IDENTIFICAÇÃO DO EMPRESÁRIO OU DA SOCIEDADE EMPRESÁRIA
'Ver8
c0000 = "|" & "0000" & "|" & "LECD" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|" & rsEmpresa!RazaoSocial & "|" & rsEmpresa!CNPJ & "|" & rsEmpresa!UF & "|" & rsEmpresa!IE & "|" & rsEmpresa!Cidade_IBGE & "|" & rsEmpresa!IM & "|" & "|" & "0" & "|" & "1" & "|" & "0" & "|" & "|" & "0" & "|" & "0" & "|" & "|" & "N" & "|" & "N" & "|" & "0" & "|" & "0" & "|" & "1" & "|"
'Ver7
'c0000 = "|" & "0000" & "|" & "LECD" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|" & rsEmpresa!RazaoSocial & "|" & rsEmpresa!CNPJ & "|" & rsEmpresa!UF & "|" & rsEmpresa!IE & "|" & rsEmpresa!Cidade_IBGE & "|" & rsEmpresa!IM & "|" & "|" & "0" & "|" & "1" & "|" & "0" & "|" & "|" & "0" & "|" & "0" & "|" & "|" & "N" & "|" & "N" & "|"

l0000 = 1
clintot = clintot + 1
clinB0 = clinB0 + 1
Print #iArq, c0000

'REGISTRO0001: ABERTURA DO BLOCO 0

'c0001
c0001 = "|" & "0001" & "|" & "0" & "|"
clintot = clintot + 1
l0001 = 1
clinB0 = clinB0 + 1
Print #iArq, c0001


'Registro0007: Outras Inscrições Cadastrais da Pessoa Jurídica
'c0007
c0007 = "|" & "0007" & "|" & "00" & "|" & "|"
clintot = clintot + 1
l0007 = 1
clinB0 = clinB0 + 1
Print #iArq, c0007

'Registro0150: Tabela de Cadastro do Participante
'c0150

'Parece que aqui somente as empresas relacionadas em
'Matriz no exterior
'Filial, inclusive agência ou dependência, no exterior
'Coligada, inclusive equiparada
'Controladora
'Controlada (exceto subsidiária integral)
'Subsidiária integral
'Controlada em conjunto
'Entidade de Propósito Específico (conforme definição da CVM)
'Participante do conglomerado, conforme norma específica do órgão regulador, exceto as que se enquadrem
'nos tipos precedentes
'Vinculadas (Art. 23 da Lei 9.430/96
'), exceto as que se enquadrem nos tipos precedentes
'Localizada em país com tributação favorecida (Art. 24 da Lei 9.430/96), exceto as que se enquadrem nos
'tipos precedentes '
GoTo pula0150:

c0150 = ""
     Do Until rsCliente.EOF = True
     cPais = ""
     cCNPJ = ""
     cCPF = ""
     'clientes
     Select Case rsCliente!Pais
     Case Is = "BRASIL"
     cPais = "01058"
     Case Else
     MsgBox ("Cliente de outro país verifique o cadastro para o código do pais, não prossiga a EFD")
     End Select
     
     Select Case Len(rsCliente!CNPJ)
     Case Is = 14
     c0150 = "|" & "0150" & "|" & rsCliente!IdCliente & "|" & rsCliente!RazaoSocial & "|" & cPais & "|" & rsCliente!CNPJ & "|" & "|" & "|" & rsCliente!UF & "|" & rsCliente!IE & "|" & "|" & rsCliente!cod_Municipio & "|" & "|" & "|"
     Case Is = 11
     c0150 = "|" & "0150" & "|" & rsCliente!IdCliente & "|" & rsCliente!RazaoSocial & "|" & cPais & "|" & "|" & rsCliente!CPF & "|" & "|" & rsCliente!UF & "|" & "|" & "|" & rsCliente!cod_Municipio & "|" & "|" & "|"
     Case Else
     c0150 = "|" & "0150" & "|" & rsCliente!IdCliente & "|" & rsCliente!RazaoSocial & "|" & cPais & "|" & rsCliente!CNPJ & "|" & "|" & "|" & rsCliente!UF & "|" & rsCliente!IE & "|" & "|" & rsCliente!cod_Municipio & "|" & "|" & "|"
     End Select
     
     
     
     Print #iArq, c0150
     l0150 = l0150 + 1
     rsCliente.MoveNext
     clintot = clintot + 1
     clinB0 = clinB0 + 1
     Loop

    'fornecedores
     
     Do Until rsFornecedor.EOF = True
     cPais = ""
     cCNPJ = ""
     cCPF = ""
     
     Select Case rsFornecedor!Pais
     Case Is = "BRASIL"
     cPais = "01058"
     Case Else
     MsgBox ("Fornecedor de outro país verifique o cadastro para o código do pais, não prossiga a EFD")
     End Select
     
     Select Case Len(rsFornecedor!CNPJ)
     Case Is = 14
     c0150 = "|" & "0150" & "|" & rsFornecedor!IdFor & "|" & rsFornecedor!RazaoSocial & "|" & cPais & "|" & rsFornecedor!CNPJ & "|" & "|" & "|" & rsFornecedor!UF & "|" & rsFornecedor!IE & "|" & "|" & rsFornecedor!cod_Municipio & "|" & "|" & "|"
     Case Is = 11
     c0150 = "|" & "0150" & "|" & rsFornecedor!IdFor & "|" & rsFornecedor!RazaoSocial & "|" & cPais & "|" & "|" & rsFornecedor!CPF & "|" & "|" & rsFornecedor!UF & "|" & "|" & "|" & rsFornecedor!cod_Municipio & "|" & "|" & "|"
     Case Else
     c0150 = "|" & "0150" & "|" & rsFornecedor!IdFor & "|" & rsFornecedor!RazaoSocial & "|" & cPais & "|" & rsFornecedor!CNPJ & "|" & "|" & "|" & rsFornecedor!UF & "|" & rsFornecedor!IE & "|" & "|" & rsFornecedor!cod_Municipio & "|" & "|" & "|"
     End Select
     
     
     Print #iArq, c0150
     rsFornecedor.MoveNext
     clintot = clintot + 1
     l0150 = l0150 + 1
     clinB0 = clinB0 + 1
     Loop

pula0150:

'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0

'c0990
clintot = clintot + 1
l0990 = l0990 + 1
clinB0 = clinB0 + 1
c0990 = "|" & "0990" & "|" & clinB0 & "|"
Print #iArq, c0990


'Bloco I: Lançamentos Contábeis
Dim cLinBI As Integer
cLinBI = 0
'Registro I001: Abertura do Bloco I
'cI001
'cI001
cI001 = "|" & "I001" & "|" & "0" & "|"
clintot = clintot + 1
lI001 = 1
cLinBI = cLinBI + 1
Print #iArq, cI001

'REGISTRO I010: IDENTIFICAÇÃO DA ESCRITURAÇÃO CONTÁBIL
'cI010
cI010 = "|" & "I010" & "|" & "G" & "|" & cLayout & "|"
clintot = clintot + 1
lI010 = 1
cLinBI = cLinBI + 1
Print #iArq, cI010



'REGISTRO I030: TERMO DE ABERTURA
'cI030
Dim cNumLivro As Long
'Select Case Year(cDtIni)
'Case Is = 2019
'cNumLivro = 1
'Case Else
'ANUAL
'cNumLivro = Year(cDtIni) - 2018 + 1
'MENSAL
cNumLivro = year(cDtIni) & month(cDtIni)
'End Select


cI030 = "|" & "I030" & "|" & "TERMO DE ABERTURA" & "|" & cNumLivro & "|" & "LIVRO DIARIO GERAL" & "|" & "8888" & "|" & rsEmpresa!RazaoSocial & "|" & rsEmpresa!NIRE & "|" & rsEmpresa!CNPJ & "|" & Replace(rsEmpresa!Dt_NIRE, "/", "") & "|" & "|" & "SOROCABA" & "|" & cSTR_DtFIM & "|"
clintot = clintot + 1
lI030 = 1
cLinBI = cLinBI + 1
Print #iArq, cI030


'RegistroI050: Plano de Contas
'cI150
'cI151
Set RSI050 = Db.OpenRecordset("SELECT * FROM EFD_I050_Contas_Contabeis")

Do Until RSI050.EOF
cI050 = "|" & "I050" & "|" & RSI050!DT_INI & "|" & RSI050!NATUREZA & "|" & RSI050!Tipo & "|" & RSI050!Nivel & "|" & RSI050!COD & "|" & RSI050!CONTA_SUPERIOR & "|" & RSI050!DESCRICAO & "|"
clintot = clintot + 1
lI050 = lI050 + 1
Print #iArq, cI050
cLinBI = cLinBI + 1

'REGISTRO I051: PLANO DE CONTAS REFERENCIAL APENAS PARA CONTA ANALITICA
If RSI050!Tipo = "A" Then
'v7
'cI051 = "|" & "I051" & "|" & "2" & "|" & "|" & RSI050!Referencial_RFB & "|"
'v8
cI051 = "|" & "I051" & "|" & "|" & RSI050!Referencial_RFB & "|"

clintot = clintot + 1
lI051 = lI051 + 1
Print #iArq, cI051
cLinBI = cLinBI + 1
Else
End If

'REGISTRO I052: CODIGOS DE AGLUTINAÇÃO
If IsNull(RSI050.Fields("Cod_Aglutinacao").Value) Then
Else
cI052 = "|" & "I052" & "|" & "|" & RSI050!Cod_Aglutinacao & "|"
clintot = clintot + 1
lI052 = lI052 + 1
Print #iArq, cI052
cLinBI = cLinBI + 1
End If


RSI050.MoveNext
Loop




'Registro I150: Saldos Periódicos
'cI150
'desativar temporariamente em debug, depois ativar

Dim cAno As Integer
Dim cMes As Integer
Dim cAnomes As String

Dim cAnoFIm As Integer
Dim cMesFim As Integer
Dim cAnoMesFim As String
Dim cDtINICIO As String
Dim cDtFinal As String
Dim cSaldoIni As Double
Dim cSaldoFim As Double

cAno = year(cDtIniVb)
cMes = Left(cDtIniVb, 2)
cAnomes = cAno & Format(cMes, "00")

cAnoFIm = year(cDtFimVb)
cMesFim = Left(cDtFimVb, 2)
cAnoMesFim = cAnoFIm & Format(cMesFim, "00")

Set rsI150 = Db.OpenRecordset("select * from EFD_I150_Calendario where anomes >=" & cAnomes & " and anomes <=" & cAnoMesFim & "")

Do Until rsI150.EOF
cI150 = "|" & "I150" & "|" & Replace(rsI150!DT_INI, "/", "") & "|" & Replace(rsI150!DT_FIM, "/", "") & "|"
clintot = clintot + 1
lI150 = lI150 + 1
cLinBI = cLinBI + 1
Print #iArq, cI150

    Set rsI155 = Db.OpenRecordset("select * from EFD_I155_Detalhe_Saldos where Data_INI = #" & Format(rsI150!DT_INI, "mm/dd/yyyy") & "# and Data_Fim = # " & Format(rsI150!DT_FIM, "mm/dd/yyyy") & " #")
    
    Do Until rsI155.EOF
    cSaldoIni = 0
    cSaldoFim = 0
    If rsI155!Saldo_Inicial < 0 Then
    cSaldoIni = rsI155!Saldo_Inicial * -1
    Else
    cSaldoIni = rsI155!Saldo_Inicial
    End If
    If rsI155!Saldo_Final < 0 Then
    cSaldoFim = rsI155!Saldo_Final * -1
    Else
    cSaldoFim = rsI155!Saldo_Final
    End If
    
    
    cI155 = "|" & "I155" & "|" & rsI155!Cod_Conta & "|" & "|" & Round(cSaldoIni, 2) & "|" & rsI155!Ind_Saldo_Ini & "|" & Round(rsI155!Total_Debitos, 2) & "|" & Round(rsI155!Total_Creditos, 2) & "|" & Round(cSaldoFim, 2) & "|" & rsI155!Ind_Saldo_Fim & "|"
    clintot = clintot + 1
    lI155 = lI155 + 1
    cLinBI = cLinBI + 1
    Print #iArq, cI155
    rsI155.MoveNext
    Loop
    
rsI150.MoveNext
Loop


'Registro I200: Lançamento Contábil
'cI200
Set rsI200 = Db.OpenRecordset("select * from EFD_I200_Lancamento_Contabil_Head where data >= #" & cDtIni & "# and data <= #" & cDtFim & "#")

Do Until rsI200.EOF
cI200 = "|" & "I200" & "|" & rsI200!ID & "|" & Replace(rsI200!Data, "/", "") & "|" & Round(rsI200!Valor, 2) & "|" & rsI200!Indicador & "|" & "|"
clintot = clintot + 1
lI200 = lI200 + 1
cLinBI = cLinBI + 1
Print #iArq, cI200
Set rsI250 = Db.OpenRecordset("select * from EFD_I200_Lancamento_Contabil where Id = " & rsI200!ID & "")
Do Until rsI250.EOF
cI250 = "|" & "I250" & "|" & rsI250!Conta & "|" & "|" & Round(rsI250!Valor, 2) & "|" & rsI250!Tipo & "|" & rsI250!Num_Nota & "|" & "|" & rsI250!Historico & "|" & "|"
clintot = clintot + 1
lI250 = lI250 + 1
cLinBI = cLinBI + 1
Print #iArq, cI250
rsI250.MoveNext
Loop
rsI200.MoveNext
Loop


'Registro I350: Saldo das Contas de Resultado Antes do Encerramento - Identificação da Data
cI350 = "|" & "I350" & "|" & cSTR_DtFIM & "|"
clintot = clintot + 1
lI350 = lI350 + 1
cLinBI = cLinBI + 1
Print #iArq, cI350

'Registro I355: Detalhes dos Saldos das Contas de Resultado Antes do Encerramento
Set rsI355 = Db.OpenRecordset("SELECT Left(ANOMES,4) AS ANO, EFD_I350_Detalhe_Saldos_Antes_Encerramento.Data_INI, EFD_I350_Detalhe_Saldos_Antes_Encerramento.Data_FIM, EFD_I350_Detalhe_Saldos_Antes_Encerramento.Cod_Conta, EFD_I350_Detalhe_Saldos_Antes_Encerramento.Saldo_Final, EFD_I350_Detalhe_Saldos_Antes_Encerramento.Ind_Saldo_Fim " & _
"FROM EFD_I350_Detalhe_Saldos_Antes_Encerramento INNER JOIN EFD_I050_Contas_Contabeis ON EFD_I350_Detalhe_Saldos_Antes_Encerramento.Cod_Conta = EFD_I050_Contas_Contabeis.COD " & _
"WHERE EFD_I350_Detalhe_Saldos_Antes_Encerramento.Data_FIM = #" & cDtFimVb & "# AND (Left(Cod_Conta,1)=3 Or Left(Cod_Conta,1)=4) AND EFD_I050_Contas_Contabeis.TIPO='A';")

Do Until rsI355.EOF

If rsI355!Saldo_Final < 0 Then
cI355 = "|" & "I355" & "|" & rsI355!Cod_Conta & "|" & "|" & Round(rsI355!Saldo_Final, 2) * -1 & "|" & rsI355!Ind_Saldo_Fim & "|"
Else
cI355 = "|" & "I355" & "|" & rsI355!Cod_Conta & "|" & "|" & Round(rsI355!Saldo_Final, 2) & "|" & rsI355!Ind_Saldo_Fim & "|"
End If

clintot = clintot + 1
lI355 = lI355 + 1
cLinBI = cLinBI + 1

Print #iArq, cI355

rsI355.MoveNext
Loop





'REGISTRO I990: Encerramento do Bloco I
'cI990
clintot = clintot + 1
lI990 = lI990 + 1
cLinBI = cLinBI + 1
cI990 = "|" & "I990" & "|" & cLinBI & "|"
Print #iArq, cI990


'BLOCO J: Demonstrações Contábeis
Dim cLinBJ As Integer
cLinBJ = 0
'Registro J001: Abertura do Bloco J
cJ001 = "|" & "J001" & "|" & "0" & "|"   '1 sem dados informados, 0 com dados informados
clintot = clintot + 1
lJ001 = 1
cLinBJ = cLinBJ + 1
Print #iArq, cJ001



'Registro J005: Demonstrações Contábeis
cJ005 = "|" & "J005" & "|" & Replace(Format(cDtIni, "dd/mm/yyyy"), "/", "") & "|" & Replace(Format(cDtFim, "dd/mm/yyyy"), "/", "") & "|" & "1" & "|" & "|"
clintot = clintot + 1
lJ005 = lJ005 + 1
cLinBJ = cLinBJ + 1
Print #iArq, cJ005

'Registro J100: Balanço Patrimonial
Dim cIndINI As String
Dim cIndFIM As String
Dim cValorIni As String
Dim cValorFim As String
Dim cPonto As Integer

Dim rsBalancoDet As Recordset
Dim rsBalanco As Recordset


Set rsBalanco = Db.OpenRecordset("select * from EFD_J100_Codigos_Aglutinacao where IND_GRP_BAL = 'A' OR IND_GRP_BAL = 'P'")
Do Until rsBalanco.EOF

Select Case rsBalanco!IND_COD_AGL
Case Is = "T"
cPonto = Len(rsBalanco!COD_AGLT) - Len(Replace(rsBalanco!COD_AGLT, ".", ""))
Set rsBalancoDet = Db.OpenRecordset("SELECT Sum(q1.saldo_inicial) AS Saldo_Inicial, Sum(q1.saldo_final) AS Saldo_Final " & _
"FROM EFD_J100_Codigos_Aglutinacao INNER JOIN ((select q1.Data_Ini, q2.Data_Fim, q2.cod_conta, q1.saldo_inicial, q2.saldo_final from " & _
"(select cod_conta, Data_Ini, Saldo_Inicial from EFD_I155_Detalhe_Saldos where data_ini = #" & cDtIni & "#) as q1 " & _
"Right Join " & _
"(select cod_conta, Data_Fim, saldo_final from EFD_I155_Detalhe_Saldos where data_fim = #" & cDtFim & "#) as q2 " & _
"on q1.cod_conta = q2.cod_conta " & _
")  AS q1 INNER JOIN EFD_I050_Contas_Contabeis ON q1.cod_conta = EFD_I050_Contas_Contabeis.COD) ON EFD_J100_Codigos_Aglutinacao.COD_AGLT = EFD_I050_Contas_Contabeis.Cod_Aglutinacao " & _
"WHERE ((Left(COD_AGLT," & Len(rsBalanco!COD_AGLT) & ") = '" & rsBalanco!COD_AGLT & "'));")


Case Is = "D"
Set rsBalancoDet = Db.OpenRecordset("SELECT EFD_J100_Codigos_Aglutinacao.COD_AGLT, EFD_J100_Codigos_Aglutinacao.DESC_COD_AGL, Sum(q1.saldo_inicial) AS Saldo_Inicial, Sum(q1.saldo_final) AS Saldo_Final " & _
"FROM EFD_J100_Codigos_Aglutinacao INNER JOIN ((select q1.Data_Ini, q2.Data_Fim, q2.cod_conta, q1.saldo_inicial, q2.saldo_final from " & _
"(select cod_conta, Data_Ini, Saldo_Inicial from EFD_I155_Detalhe_Saldos where data_ini = #" & cDtIni & "#) as q1 " & _
"Right Join " & _
"(select cod_conta, Data_Fim, saldo_final from EFD_I155_Detalhe_Saldos where data_fim = #" & cDtFim & "#) as q2 " & _
"on q1.cod_conta = q2.cod_conta " & _
")  AS q1 INNER JOIN EFD_I050_Contas_Contabeis ON q1.cod_conta = EFD_I050_Contas_Contabeis.COD) ON EFD_J100_Codigos_Aglutinacao.COD_AGLT = EFD_I050_Contas_Contabeis.Cod_Aglutinacao " & _
"GROUP BY EFD_J100_Codigos_Aglutinacao.COD_AGLT, EFD_J100_Codigos_Aglutinacao.DESC_COD_AGL " & _
"HAVING (((EFD_J100_Codigos_Aglutinacao.COD_AGLT) = '" & rsBalanco!COD_AGLT & "'))")
End Select


cIndINI = ""
cIndFIM = ""
cValorIni = 0
cValorFim = 0

'VALOR INI
If rsBalancoDet.EOF = True And rsBalancoDet.BOF = True Then
GoTo PulaBalanco:
Else
End If

If IsNull(rsBalancoDet.Fields("Saldo_Inicial").Value) Then
cIndINI = "C"
cValorIni = 0
Else
If rsBalancoDet!Saldo_Inicial < 0 Then
cIndINI = "D"
cValorIni = rsBalancoDet!Saldo_Inicial * -1
    'se for depreciação
    If rsBalanco!COD_AGLT = "1.2.5" Then
    cIndINI = "C"
    cValorIni = rsBalancoDet!Saldo_Inicial
    'se for depreciação
    Else
    End If

Else
cIndINI = "C"
cValorIni = rsBalancoDet!Saldo_Inicial
End If
End If


'VALOR FIM
If IsNull(rsBalancoDet.Fields("Saldo_Final").Value) Then
cIndFIM = "C"
cValorFim = 0
Else
If rsBalancoDet!Saldo_Final < 0 Then
cIndFIM = "D"
cValorFim = rsBalancoDet!Saldo_Final * -1
    'se for depreciação
    If rsBalanco!COD_AGLT = "1.2.5" Then
    cIndINI = "C"
    cValorIni = rsBalancoDet!Saldo_Final
    'se for depreciação
    Else
    End If

Else
cIndFIM = "C"
cValorFim = rsBalancoDet!Saldo_Final
End If
End If



If IsNull(rsBalanco.Fields("COD_AGL_SUP").Value) Then
cJ100 = "|" & "J100" & "|" & rsBalanco!COD_AGLT & "|" & rsBalanco!IND_COD_AGL & "|" & rsBalanco!NIVEL_AGL & "|" & "|" & rsBalanco!IND_GRP_BAL & "|" & rsBalanco!DESC_COD_AGL & "|" & Replace(Round(cValorIni, 2), ".", ",") & "|" & cIndINI & "|" & Replace(Round(cValorFim, 2), ".", ",") & "|" & cIndFIM & "|" & "|"
Else
cJ100 = "|" & "J100" & "|" & rsBalanco!COD_AGLT & "|" & rsBalanco!IND_COD_AGL & "|" & rsBalanco!NIVEL_AGL & "|" & rsBalanco!COD_AGL_SUP & "|" & rsBalanco!IND_GRP_BAL & "|" & rsBalanco!DESC_COD_AGL & "|" & Replace(Round(cValorIni, 2), ".", ",") & "|" & cIndINI & "|" & Replace(Round(cValorFim, 2), ".", ",") & "|" & cIndFIM & "|" & "|"
End If
clintot = clintot + 1
lJ100 = lJ100 + 1
cLinBJ = cLinBJ + 1
Print #iArq, cJ100

PulaBalanco:
rsBalanco.MoveNext
Loop

'Registro J150: Demonstração do Resultado do Exercício (DRE)


'limpa valtemp
strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp = 0;")
Conn.Execute strSQL


Set rsBalanco = Db.OpenRecordset("select * from EFD_J100_Codigos_Aglutinacao WHERE (IND_GRP_BAL = 'R' and IND_COD_AGL = 'D') OR (IND_GRP_BAL = 'D' and IND_COD_AGL = 'D') order by ordenacao asc")
'DoCmd.setwarnings (True)
'TOTALIZADORES SALDO
Dim cValINI As Double
Dim cValFIM As Double
Dim cFlagT As Boolean


'APENAS TIPO 'D' - Detalhe
Do Until rsBalanco.EOF
      
        Set rsBalancoDet = Db.OpenRecordset("SELECT sum(q1.saldo_inicial) as Saldo_Inicial, Sum(q1.saldo_final) AS Saldo_final " & _
        "FROM EFD_J100_Codigos_Aglutinacao INNER JOIN ((SELECT q1.Data_Fim, q1.cod_conta, q1.saldo_inicial, q1.saldo_final " & _
        "FROM (SELECT cod_conta, Data_Fim, Saldo_Inicial, Saldo_final FROM EFD_I350_Detalhe_Saldos_ANTES_Encerramento WHERE data_fim = #" & cDtFim & "#)  AS q1)  AS q1 INNER JOIN EFD_I050_Contas_Contabeis ON q1.cod_conta = EFD_I050_Contas_Contabeis.COD) ON EFD_J100_Codigos_Aglutinacao.COD_AGLT = EFD_I050_Contas_Contabeis.Cod_Aglutinacao " & _
        "WHERE COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
        'eom = DateAdd("d", -1, DateAdd("m", 1, DateSerial(Year(input_date), Month(input_date), 1)))
        cDataAnt = Format((DateAdd("m", -1, cDtFim)), "mm/dd/yyyy")
        cDataAnt = DateAdd("d", -1, DateAdd("m", 1, DateSerial(year(cDataAnt), month(cDataAnt), 1)))
        Set rsBalancoDet_Ant = Db.OpenRecordset("SELECT sum(q1.saldo_inicial) as Saldo_Inicial, Sum(q1.saldo_final) AS Saldo_final " & _
        "FROM EFD_J100_Codigos_Aglutinacao INNER JOIN ((SELECT q1.Data_Fim, q1.cod_conta, q1.saldo_inicial, q1.saldo_final " & _
        "FROM (SELECT cod_conta, Data_Fim, Saldo_Inicial, Saldo_final FROM EFD_I350_Detalhe_Saldos_ANTES_Encerramento WHERE data_fim = #" & cDataAnt & "#)  AS q1)  AS q1 INNER JOIN EFD_I050_Contas_Contabeis ON q1.cod_conta = EFD_I050_Contas_Contabeis.COD) ON EFD_J100_Codigos_Aglutinacao.COD_AGLT = EFD_I050_Contas_Contabeis.Cod_Aglutinacao " & _
        "WHERE COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
        
          
          If IsNull(rsBalancoDet.Fields("Saldo_Final").Value) Then
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp = 0 where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          Else
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp = " & Replace(rsBalancoDet!Saldo_Final, ",", ".") & " where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          End If
          Conn.Execute strSQL
          
          If IsNull(rsBalancoDet_Ant.Fields("Saldo_Final").Value) Then
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp_Ant = 0 where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          Else
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp_Ant = " & Replace(rsBalancoDet_Ant!Saldo_Final, ",", ".") & " where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          End If
          Conn.Execute strSQL
           
rsBalanco.MoveNext
Loop
        

'APENAS TIPO 'T' - Totais
Set rsBalanco = Db.OpenRecordset("select * from EFD_J100_Codigos_Aglutinacao WHERE (IND_GRP_BAL = 'R' and IND_COD_AGL = 'T') OR (IND_GRP_BAL = 'D' and IND_COD_AGL = 'T') order by ordenacao asc")

Do Until rsBalanco.EOF
        Set rsBalancoDet = Db.OpenRecordset("SELECT Sum(ValTemp_Ant) as Saldo_Inicial, Sum(ValTemp) AS Saldo_Final " & _
        "FROM EFD_J100_Codigos_Aglutinacao " & _
        "WHERE COD_AGL_SUP = '" & rsBalanco!COD_AGLT & "';")
               
          If IsNull(rsBalancoDet.Fields("Saldo_Final").Value) Then
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp = 0 where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          Else
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp = " & Replace(rsBalancoDet!Saldo_Final, ",", ".") & " where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          End If
          Conn.Execute strSQL
          
          If IsNull(rsBalancoDet.Fields("Saldo_Inicial").Value) Then
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp_Ant = 0 where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          Else
          strSQL = ("update EFD_J100_Codigos_Aglutinacao set ValTemp_Ant = " & Replace(rsBalancoDet!Saldo_Inicial, ",", ".") & " where COD_AGLT = '" & rsBalanco!COD_AGLT & "';")
          End If
          Conn.Execute strSQL
          
rsBalanco.MoveNext
Loop
        
        
'LANÇA DADOS NO TXT E FAZ TRATAMETO
Set rsBalanco = Db.OpenRecordset("select * from EFD_J100_Codigos_Aglutinacao WHERE IND_GRP_BAL = 'R' OR IND_GRP_BAL = 'D' order by ordenacao asc")
        
Do Until rsBalanco.EOF
               
        cValINI = rsBalanco!ValTemp_Ant
        cValFIM = rsBalanco!ValTemp
        
        If cValINI < 0 Then
        cIndINI = "D"
        cValINI = cValINI * -1
        Else
        cIndINI = "C"
        cValINI = cValINI
        End If
        
        
        If cValFIM < 0 Then
        cIndFIM = "D"
        cValFIM = cValFIM * -1
        Else
        cIndFIM = "C"
        cValFIM = cValFIM
        End If
             
        'v7
        'cJ150 = "|" & "J150" & "|" & rsBalanco!COD_AGLT & "|" & rsBalanco!IND_COD_AGL & "|" & rsBalanco!NIVEL_AGL & "|" & rsBalanco!COD_AGL_SUP & "|" & rsBalanco!DESC_COD_AGL & "|" & Round(cValFIM, 2) & "|" & cIndFIM & "|" & rsBalanco!IND_GRP_BAL & "|" & "|"
        'v8
        cJ150 = "|" & "J150" & "|" & rsBalanco!Ordenacao & "|" & rsBalanco!COD_AGLT & "|" & rsBalanco!IND_COD_AGL & "|" & rsBalanco!NIVEL_AGL & "|" & rsBalanco!COD_AGL_SUP & "|" & rsBalanco!DESC_COD_AGL & "|" & Round(cValINI, 2) & "|" & cIndINI & "|" & Round(cValFIM, 2) & "|" & cIndFIM & "|" & rsBalanco!IND_GRP_BAL & "|" & "|"
        
      
 

clintot = clintot + 1
lJ150 = lJ150 + 1
cLinBJ = cLinBJ + 1
Print #iArq, cJ150

cFlagT = False
cValINI = 0
cValFIM = 0
rsBalanco.MoveNext
Loop




'Registro J210: DLPA - Demonstração de Lucros ou Prejuízos Acumulados/DMPL - Demonstração de Mutações do Patrimônio Líquido


'Registro J900: Termo de Encerramento
clintot = clintot + 1
lJ900 = 1
cLinBJ = cLinBJ + 1
cJ900 = "|" & "J900" & "|" & "TERMO DE ENCERRAMENTO" & "|" & cNumLivro & "|" & "LIVRO DIARIO GERAL" & "|" & rsEmpresa!RazaoSocial & "|" & "8888" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|"
Print #iArq, cJ900


'Registro J930: Signatários da Escrituração
Set rsSignatario = Db.OpenRecordset("tb_Signatarios_ECD")
lJ930 = 0
Do Until rsSignatario.EOF
cJ930 = "|" & "J930" & "|" & rsSignatario!nome & "|" & rsSignatario!CPF_CNPJ & "|" & rsSignatario!Qualificacao & "|" & rsSignatario!cod_qualificacao & "|" & rsSignatario!CRC_Contabilista & "|" & rsSignatario!Email & "|" & rsSignatario!Fone & "|" & rsSignatario!UF_CRC & "|" & rsSignatario!NUM_SEQ_CRC & "|" & rsSignatario!DT_CRC & "|" & rsSignatario!RESP_LEGAL & "|"
Print #iArq, cJ930
clintot = clintot + 1
cLinBJ = cLinBJ + 1
lJ930 = lJ930 + 1
rsSignatario.MoveNext
Loop

'Registro J990: Encerramento do Bloco J
clintot = clintot + 1
lJ990 = 1
cLinBJ = cLinBJ + 1
cJ990 = "|" & "J990" & "|" & cLinBJ & "|"
Print #iArq, cJ990

'Registro K001: Abertura do Bloco K
'cK001 = "|" & "K001" & "|" & "1" & "|"
'Print #iArq, cK001
'clintot = clintot + 1
'lK001 = 1
'cLinBK = cLinBK + 1

'Registro K990: Encerramento do Bloco K
'clintot = clintot + 1
'lK990 = 1
'cLinBK = cLinBK + 1
'cK990 = "|" & "K990" & "|" & cLinBK & "|"
'Print #iArq, cK990



'BLOCO 9: Controle e Encerramento do Arquivo Digital
'Registro 9001: Abertura do Bloco 9
Dim cLin9 As Integer
cLin9 = 0

'c9001
c9001 = "|" & "9001" & "|" & "0" & "|"
Print #iArq, c9001
cLin9 = cLin9 + 1
l9001 = l9001 + 1
l99 = l99 + 1

'Registro 9900: Registros do Arquivo



'BLOCO 0:
'c0000
c9900 = "|" & "9900" & "|" & "0000" & "|" & l0000 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0001
c9900 = "|" & "9900" & "|" & "0001" & "|" & l0001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0007
c9900 = "|" & "9900" & "|" & "0007" & "|" & l0007 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0150
'c9900 = "|" & "9900" & "|" & "0150" & "|" & l0150 & "|"
'Print #iArq, c9900
'l99 = l99 + 1
'c0990
c9900 = "|" & "9900" & "|" & "0990" & "|" & l0990 & "|"
Print #iArq, c9900
l99 = l99 + 1
''BLOCO I:
'cI001
c9900 = "|" & "9900" & "|" & "I001" & "|" & lI001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI010
c9900 = "|" & "9900" & "|" & "I010" & "|" & lI010 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI030
c9900 = "|" & "9900" & "|" & "I030" & "|" & lI030 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI050
c9900 = "|" & "9900" & "|" & "I050" & "|" & lI050 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI051
c9900 = "|" & "9900" & "|" & "I051" & "|" & lI051 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI052
c9900 = "|" & "9900" & "|" & "I052" & "|" & lI052 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI150
c9900 = "|" & "9900" & "|" & "I150" & "|" & lI150 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI151
'c9900 = "|" & "9900" & "|" & "I151" & "|" & lI151 & "|"
'Print #iArq, c9900
'l99 = l99 + 1
'cI155
c9900 = "|" & "9900" & "|" & "I155" & "|" & lI155 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI200
c9900 = "|" & "9900" & "|" & "I200" & "|" & lI200 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI250
c9900 = "|" & "9900" & "|" & "I250" & "|" & lI250 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI350
c9900 = "|" & "9900" & "|" & "I350" & "|" & lI350 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI355
c9900 = "|" & "9900" & "|" & "I355" & "|" & lI355 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cI990
c9900 = "|" & "9900" & "|" & "I990" & "|" & lI990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO J:
'cJ001
c9900 = "|" & "9900" & "|" & "J001" & "|" & lJ001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cJ005
c9900 = "|" & "9900" & "|" & "J005" & "|" & lJ005 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cJ100
c9900 = "|" & "9900" & "|" & "J100" & "|" & lJ100 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cJ150
c9900 = "|" & "9900" & "|" & "J150" & "|" & lJ150 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cJ900
c9900 = "|" & "9900" & "|" & "J900" & "|" & lJ900 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cJ930
c9900 = "|" & "9900" & "|" & "J930" & "|" & lJ930 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cJ990
c9900 = "|" & "9900" & "|" & "J990" & "|" & lJ990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO K
'cK001
'c9900 = "|" & "9900" & "|" & "K001" & "|" & lK001 & "|"
'Print #iArq, c9900
'l99 = l99 + 1
'cK990
'c9900 = "|" & "9900" & "|" & "K990" & "|" & lK990 & "|"
'Print #iArq, c9900
'l99 = l99 + 1



'Registro 9990: Encerramento do Bloco 9
'9001
c9900 = "|" & "9900" & "|" & "9001" & "|" & l9001 & "|"
Print #iArq, c9900
l99 = l99 + 1





'Registro 9999: Encerramento do Arquivo Digital
'9990
'c9001 = "|" & "9900" & "|" & "9990" & "|" & "0" & "|"
'Print #iArq, c9001
'9999
'c9001 = "|" & "9999" & "|" & "9999" & "|" & "1" & "|"
'Print #iArq, c9001
l99 = l99 + 1
l99 = l99 + 1
'totalizador 9900
c9900 = "|" & "9900" & "|" & "9900" & "|" & l99 & "|"
Print #iArq, c9900


'totalizador 9900
c9900 = "|" & "9900" & "|" & "9990" & "|" & "1" & "|"
Print #iArq, c9900


'totalizador 9900
c9900 = "|" & "9900" & "|" & "9999" & "|" & "1" & "|"
Print #iArq, c9900

'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9

l99 = l99 + 1
l99 = l99 + 1
l99 = l99 + 1
c9990 = "|" & "9990" & "|" & l99 & "|"
Print #iArq, c9990



cTotal9999 = l0000 + l0001 + l0007 + l0150 + l0990 + lI001 + lI010 + lI030 + lI050 + lI051 + lI052 + lI150 + lI155 + lI200 + lI250 + lI350 + lI355 + lI990 + lJ001 + lJ005 + lJ100 + lJ150 + lJ900 + lJ930 + lJ990 + lK001 + lK990 + l9001 + l9900
c9999 = "|" & "9999" & "|" & cTotal9999 + l99 - 1 & "|"
Print #iArq, c9999


Close #iArq
'DoCmd.setwarnings (True)
MsgBox ("Arquivo Gerado. Linhas: " & cTotal9999 + l99 - 1 & "")

Call TextFile_FindReplace(cPath, "8888", cTotal9999 + l99 - 1)


End Function



Public Sub TextFile_FindReplace(path As String, str_look As String, str_updt As String)
'PURPOSE: Modify Contents of a text file using Find/Replace
'SOURCE: www.TheSpreadsheetGuru.com

Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String

'File Path of Text File
  FilePath = path

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file in a Read State
  Open FilePath For Input As TextFile

'Store file content inside a variable
  FileContent = Input(LOF(TextFile) - 1, TextFile)

'Clost Text File
  Close TextFile
  
'Find/Replace
  FileContent = Replace(FileContent, str_look, str_updt)

'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile

'Open the text file in a Write State
  Open FilePath For Output As TextFile
  
'Write New Text data to file
  
  Print #TextFile, FileContent
    
  
'Close Text File
  Close TextFile
  
    'Delete last line
    'Dim lines As New List(Of String)(IO.File.ReadAllLines(LocalAppData & "\settings.txt"))
    'Remove the line to delete, e.g.
    'lines.RemoveAt (lines.Count - 1)
    'IO.File.WriteAllLines(LocalAppData & "\settings.txt", lines.ToArray())
    
    

End Sub



