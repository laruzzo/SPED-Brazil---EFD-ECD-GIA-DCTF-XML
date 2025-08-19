Attribute VB_Name = "modGeraArquivo_EFD_Contribuicoes"
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

Public Function Gerar_EFD_Contribuicoes(cDtIni As String, cDtFim As String, clocal As String, cDtINI_Contabil As String, cIDIventario As String)
Call ConnectToDataBase

'ARQUIVO EFD Contribuicoes
'DoCmd.setwarnings (False)

'EXPORTAR ARQUIVO TXT
Dim iArq As Long
iArq = FreeFile

Open clocal & "\EFD_CONTR_" & month(cDtIni) & "_" & year(cDtIni) & ".txt" For Output As iArq

'Print #iArq, c0000 & Chr(13); c0001 & Chr(13); c0005 & Chr(13); c0100 & Chr(13); c0150 & Chr(13) & c0190 & Chr(13) & c0200 & Chr(13) & c0300 & Chr(13) & c0305 & Chr(13) & c0400 & Chr(13) & c0500 & Chr(13); c0600 & Chr(13) & c0990 & Chr(13) & cC001 & Chr(13) & cC100 & Chr(13) & cC170 & Chr(13) & cC190 & Chr(13) & cC500 & Chr(13) & cC501 & Chr(13) & cC990 & Chr(13) & cD001 & Chr(13) & cD190 & Chr(13) & cD990 & Chr(13) & cE001 & Chr(13) & cE100
'Print #iArq, c0000
Dim cSTR_DtINI As String
Dim cSTR_DtFIM As String

cSTR_DtINI = Replace(Format(cDtIni, "dd/mm/yyyy"), "/", "")
cSTR_DtFIM = Replace(Format(cDtFim, "dd/mm/yyyy"), "/", "")


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
Set rsCliente = Db.OpenRecordset("SELECT tbCliente.IDCliente, tbCliente.Tipo, tbCliente.Cnpj, tbCliente.RazaoSocial, tbCliente.IE, tbCliente.CRT, tbCliente.CEP, tbCliente.Logradouro, tbCliente.Nro, tbCliente.Compl, tbCliente.Bairro, tbCliente.UF, tbCliente.cod_Municipio, tbCliente.Municipio, tbCliente.Pais, tbCliente.Fone, tbCliente.Email " & _
"FROM tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente " & _
"WHERE (((tbVendas.DataEmissao)>=# " & cDtIniVb & "  # And (tbVendas.DataEmissao)<=# " & cDtFimVb & " #)) " & _
"GROUP BY tbCliente.IDCliente, tbCliente.Tipo, tbCliente.Cnpj, tbCliente.RazaoSocial, tbCliente.IE, tbCliente.CRT, tbCliente.CEP, tbCliente.Logradouro, tbCliente.Nro, tbCliente.Compl, tbCliente.Bairro, tbCliente.UF, tbCliente.cod_Municipio, tbCliente.Municipio, tbCliente.Pais, tbCliente.Fone, tbCliente.Email;")


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
strSQL = ("delete from tbCadProd_Ativo_temp")
Conn.Execute strSQL

'vendas
strSQL = ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
"SELECT tbCadProd.IDProd " & _
"FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= '" & cDtIni & "' And (tbVendas.DataEmissao) <= '" & cDtFim & " 23:59:59" & "')) " & _
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
"FROM tbVendas, tbCadProd INNER JOIN tbImobilizado ON tbCadProd.IDProd = tbImobilizado.IDProd " & _
"LEFT OUTER JOIN  tbCadProd_Ativo_temp ON tbCadProd.IDProd = tbCadProd_Ativo_temp.IDProd " & _
"WHERE tbImobilizado.DataEmissao >= '" & cDtIni & "' And tbImobilizado.DataEmissao <= '" & cDtFim & "' and  tbCadProd_Ativo_temp.IDProd is null " & _
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





'após 01/06/2018 ver 4
'após 01/01/2019 ver 5
'após 01/01/2020 ver 6
cCodVer = "006"



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


'BLOCO 0: ABERTURA, IDENTIFICAÇÃO E REFERÊNCIAS.
'Registro 0000: Abertura do Arquivo Digital e Identificação da Pessoa Jurídica
'Registro 0100: Dados do Contabilista
'Registro 0110: Regimes de Apuração da Contribuição Social e de Apropriação de Crédito
'Registro 0111:Tabela de Receita Bruta Mensal Para Fins de Rateio de Créditos Comuns
'Registro 0140: Tabela de Cadastro de Estabelecimentos
'Registro 0208: Código de Grupos por
'Marca Comercial–Refri (bebidas frias).
'Registro 0400: Tabela de Natureza da Operação/Prestação
'Registro 0600: Centro de Custos



'c0000
c0000 = "|" & "0000" & "|" & cCodVer & "|" & 0 & "|" & "|" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|" & rsEmpresa!RazaoSocial & "|" & rsEmpresa!CNPJ & "|" & rsEmpresa!UF & "|" & rsEmpresa!Cidade_IBGE & "|" & "|" & "00" & "|" & "0" & "|"
l0000 = 1
clintot = clintot + 1
Print #iArq, c0000

'c0001
c0001 = "|" & "0001" & "|" & "0" & "|"
clintot = clintot + 1
l0001 = 1
Print #iArq, c0001

'c0035 Omitido

'c0100
c0100 = "|" & "0100" & "|" & rsContador!NomeContador & "|" & rsContador!CPFContador & "|" & rsContador!CRCContador & "|" & rsContador!CNPJEscritorio & "|" & rsContador!CEPEscritorio & "|" & rsContador!ENDEscritorio & "|" & rsContador!NumeroEscritorio & "|" & rsContador!ComplEscritorio & "|" & rsContador!BairroEscritorio & "|" & rsContador!TelefoneEscritorio & "||" & rsContador!EmailEscritorio & "|" & rsContador!CodMunicipioEscritorio & "|"
clintot = clintot + 1
l0100 = 1
Print #iArq, c0100

'Não tem o registro 0005 estranho
'c0005 = "|" & "0005" & "|" & rsEmpresa!NomeFantasia & "|" & rsEmpresa!CEP & "|" & rsEmpresa!Logradouro & "|" & rsEmpresa!Num & "|" & rsEmpresa!Compl & "|" & rsEmpresa!Bairro & "|" & rsEmpresa!Fone & "||" & rsEmpresa!Email & "|"
'clintot = clintot + 1
'l0005 = 1
'Print #iArq, c0005


'c0110
'Cervejaria é Regime diferenciado, aliquota reduzida e aliquota concentrada
c0110 = "|" & "0110" & "|" & "3" & "|" & "1" & "|" & "2" & "|" & "|"
clintot = clintot + 1
l0110 = 1
Print #iArq, c0110

'c0111
'Omitido

'c0120
'Omitido

'Registro 0140: Tabela de Cadastro de Estabelecimentos
c0140 = "|" & "0140" & "|" & "F001" & "|" & rsEmpresa!RazaoSocial & "|" & rsEmpresa!CNPJ & "|" & rsEmpresa!UF & "|" & rsEmpresa!IE & "|" & rsEmpresa!Cidade_IBGE & "|" & rsEmpresa!IM & "|" & "|"
clintot = clintot + 1
l0140 = 1
Print #iArq, c0140


'Registro 0145: Regime de Apuração da Contribuição Previdenciária Sobre a Receita Bruta
'c0145 Omitido

'Registro 0150: Tabela de Cadastro do Participante
'c0150
c0150 = ""
     Do Until rsCliente.EOF = True
       
       If rsCliente!CRT = "CONSUMIDOR" Then
      
     Else
     
     
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
     c0150 = "|" & "0150" & "|" & rsCliente!IdCliente & "|" & rsCliente!RazaoSocial & "|" & cPais & "|" & rsCliente!CNPJ & "|" & "|" & rsCliente!IE & "|" & rsCliente!cod_Municipio & "|" & "|" & rsCliente!Logradouro & "|" & rsCliente!Nro & "|" & rsCliente!compl & "|" & rsCliente!Bairro & "|"
     Case Is = 11
     c0150 = "|" & "0150" & "|" & rsCliente!IdCliente & "|" & rsCliente!RazaoSocial & "|" & cPais & "|" & "|" & rsCliente!CNPJ & "|" & rsCliente!IE & "|" & rsCliente!cod_Municipio & "|" & "|" & rsCliente!Logradouro & "|" & rsCliente!Nro & "|" & rsCliente!compl & "|" & rsCliente!Bairro & "|"
     Case Else
     c0150 = "|" & "0150" & "|" & rsCliente!IdCliente & "|" & rsCliente!RazaoSocial & "|" & cPais & "|" & rsCliente!CNPJ & "|" & "|" & rsCliente!IE & "|" & rsCliente!cod_Municipio & "|" & "|" & rsCliente!Logradouro & "|" & rsCliente!Nro & "|" & rsCliente!compl & "|" & rsCliente!Bairro & "|"
     End Select
     
     
     
     Print #iArq, c0150
     l0150 = l0150 + 1
     
     clintot = clintot + 1
     
     End If
     rsCliente.MoveNext
     
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
     c0150 = "|" & "0150" & "|" & rsFornecedor!IdFor & "|" & rsFornecedor!RazaoSocial & "|" & cPais & "|" & rsFornecedor!CNPJ & "|" & "|" & rsFornecedor!IE & "|" & rsFornecedor!cod_Municipio & "|" & "|" & rsFornecedor!Logradouro & "|" & rsFornecedor!Nro & "|" & rsFornecedor!compl & "|" & rsFornecedor!Bairro & "|"
     Case Is = 11
     c0150 = "|" & "0150" & "|" & rsFornecedor!IdFor & "|" & rsFornecedor!RazaoSocial & "|" & cPais & "|" & "|" & rsFornecedor!CNPJ & "|" & rsFornecedor!IE & "|" & rsFornecedor!cod_Municipio & "|" & "|" & rsFornecedor!Logradouro & "|" & rsFornecedor!Nro & "|" & rsFornecedor!compl & "|" & rsFornecedor!Bairro & "|"
     Case Else
     c0150 = "|" & "0150" & "|" & rsFornecedor!IdFor & "|" & rsFornecedor!RazaoSocial & "|" & cPais & "|" & rsFornecedor!CNPJ & "|" & "|" & rsFornecedor!IE & "|" & rsFornecedor!cod_Municipio & "|" & "|" & rsFornecedor!Logradouro & "|" & rsFornecedor!Nro & "|" & rsFornecedor!compl & "|" & rsFornecedor!Bairro & "|"
     End Select
     
     
     Print #iArq, c0150
     rsFornecedor.MoveNext
     clintot = clintot + 1
     l0150 = l0150 + 1
     Loop
     
 'c0190
 c0190 = ""
 Do Until rsMedidas.EOF = True
 If rsMedidas!Unid = "KW" Then
 rsMedidas.MoveNext
 Else
 c0190 = "|" & "0190" & "|" & rsMedidas!Unid & "|" & rsMedidas!DESC & "|"
 Print #iArq, c0190
 rsMedidas.MoveNext
 clintot = clintot + 1
 l0190 = l0190 + 1
 End If
 Loop
 

'c0200
 c0200 = ""
 Do Until rsCadProd.EOF = True
 cTipoItem = ""
 
       cTipoItem = "99"
       If rsCadProd!revenda = "SIM" Then
       cTipoItem = "00"
       Else: End If
       If rsCadProd!MAT_PRIMA = "SIM" Then
       cTipoItem = "01"
       Else: End If
       If rsCadProd!EMBALAGEM = "SIM" Then
       cTipoItem = "02"
       Else: End If
       If rsCadProd!PROD_FINAL = "SIM" Then
       cTipoItem = "04"
       Else: End If
       If rsCadProd!CONSUMO = "SIM" Then
       cTipoItem = "07"
       Else: End If
       If rsCadProd!IMOBILIZADO = "SIM" Then
       cTipoItem = "08"
       Else: End If
 If rsCadProd!IDProd = 2800 Then
 rsCadProd.MoveNext
 Else
 c0200 = "|" & "0200" & "|" & rsCadProd!IDProd & "|" & rsCadProd!DescProd & "|" & rsCadProd!EAN & "|" & "|" & rsCadProd!Unid & "|" & cTipoItem & "|" & rsCadProd!NCM & "|" & "|" & "|" & "|" & "|"
 Print #iArq, c0200
 rsCadProd.MoveNext
 l0200 = l0200 + 1
 clintot = clintot + 1
 End If
 Loop
 

'c0208
'Tabela XII em garrafa
'Demais Marcas Nacionais Especiais - 2
'2203.00.00 e 2202.90.00 Ex 03
'Vidro Descartável e outras embalagens não especificadas
'Cervejas de malte e cervejas sem álcool
'10. Para efeito de cálculo dos tributos da Tabela XII,o valor base representa 35,7%
'(trinta e cinco inteiros e sete décimos por cento) do preço de referência

'Tabela XIII chopp em barril
'2203.00.00 Ex 01
'Neste sentido, o registro 0208 não precisa mais ser escriturado, para os fatos geradores a partir de maio de 2015.


'c0400
Dim c0400Data
c0400Data = "NAO"

c0400 = ""
Do Until rsCFOP.EOF
If rsCFOP!CFOP_ESCRITURADA = 1252 Then
rsCFOP.MoveNext
Else
c0400 = "|" & "0400" & "|" & rsCFOP!CFOP_ESCRITURADA & "|" & rsCFOP!CFOP_ESC_DESC & "|"
Print #iArq, c0400
c0400Data = "SIM"
rsCFOP.MoveNext
clintot = clintot + 1
l0400 = l0400 + 1
End If
Loop

'c0450
'Omitido

'c0500
c0500 = ""
Do Until rsContas.EOF = True
c0500 = "|" & "0500" & "|" & cSTR_DtINI & "|" & rsContas!Cod_Natureza & "|" & rsContas!Cod_Indicador & "|" & "1" & "|" & rsContas!ID & "|" & rsContas!Desc_CodNatureza & "|" & "|" & "|"
Print #iArq, c0500

rsContas.MoveNext
clintot = clintot + 1
l0500 = l0500 + 1
Loop


'c0600
c0600 = "|" & "0600" & "|" & cSTR_DtINI & "|" & "3" & "|" & "área produtiva" & "|"
Print #iArq, c0600
clintot = clintot + 1
l0600 = l0600 + 1
c0600 = "|" & "0600" & "|" & cSTR_DtINI & "|" & "5" & "|" & "área administrativa" & "|"
Print #iArq, c0600
clintot = clintot + 1
l0600 = l0600 + 1

'c0990
clintot = clintot + 1
l0990 = l0990 + 1
c0990 = "|" & "0990" & "|" & clintot & "|"
Print #iArq, c0990


'BLOCO A: Documentos Fiscais-Serviços (Sujeitos ao ISS)
'Registro A001: Abertura do Bloco A
'cA001
cA001 = "|" & "A001" & "|" & "1" & "|"
Print #iArq, cA001
lA001 = lA001 + 1
cLinA = cLinA + 1


'Registro A990:Encerramento do Bloco A
'cA990
cLinA = cLinA + 1
lA990 = lA990 + 1
cA990 = "|" & "A990" & "|" & cLinA & "|"


'BLOCO C: Documentos Fiscais–I-Mercadorias (ICMS / IPI)
'Registro C001: Abertura do Bloco C
'cC001

Dim cC001 As String
Dim cC100 As String


Dim cLinC As Integer


Dim rsCompra As DAO.Recordset
Set rsCompra = Db.OpenRecordset("select * from tbCompras where dataemissao >= #" & cDtIniVb & "# and dataemissao <= #" & cDtFimVb & "# and IdFornecedor <> 1131 ")

Dim rsVenda As DAO.Recordset
Set rsVenda = Db.OpenRecordset("select * from tbVendas where dataemissao >= #" & cDtIniVb & "# and dataemissao <= #" & cDtFimVb & "# and TipoNF ='1-SAIDA' and [STATUS] = 'ATIVO' and NatOperacao <> 'Venda Cupom Fiscal SAT'")

'Dim rsTransp As DAO.Recordset
'Set rsTransp = DB.OpenRecordset("select * from tbTransportes where dataemissao >= #" & cDtINI & "# and dataemissao <= #" & cDtFIM & "# and tomador = 'REMETENTE' AND RemetenteCNPJ = '23866944000141'")

Dim rsCompraDet As DAO.Recordset
'Set rsCompraDet = db.OpenRecordset("SELECT tbCompras.ID as ID, tbComprasDet.CST_ICMS, tbComprasDet.CST_IPI, tbComprasDet.CST_PIS, tbComprasDet.CST_Cofins, tbComprasDet.ID as ID_DET,tbComprasDet.IDCompra as ID_Compra,  tbComprasDet.IDProd, tbComprasDet.Qnt, tbCadProd.Unid, tbCompras.dataemissao, tbComprasDet.ValorTot, tbComprasDet.VlrDesc, tbComprasDet.CST, tbComprasDet.CFOP, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.BaseCalculo, tbComprasDet.Aliq_ICMS, tbComprasDet.Valor_ICMS, tbComprasDet.BaseCalc_ST, tbComprasDet.Aliq_ICMS_ST, tbComprasDet.Valor_ICMS_ST, tbComprasDet.Aliq_IPI, tbComprasDet.Valor_IPI, tbComprasDet.Aliq_PIS, tbComprasDet.Aliq_Cofins, tbComprasDet.Valor_PIS, tbComprasDet.Valor_Cofins " & _
'"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
'"WHERE dataemissao >= #" & cDtINI & "# and dataemissao <= #" & cDtFIM & "# ORDER BY tbCompras.ID, tbComprasDet.ID;")

Dim rsVendaDet As DAO.Recordset
'Set rsVendaDet = db.OpenRecordset("SELECT tbVendas.ID as ID, tbVendasDet.ID as ID_DET,tbVendasDet.IDVenda as ID_Venda,  tbVendasDet.IDProd, tbVendasDet.Qnt, tbCadProd.Unid, tbVendas.dataemissao, tbVendasDet.ValorTot, tbVendasDet.VlrDesc, tbVendasDet.CST, tbVendasDet.CFOP, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.BaseCalculo, tbVendasDet.Aliq_ICMS, tbVendasDet.Valor_ICMS, tbVendasDet.BaseCalc_ST, tbVendasDet.Aliq_ICMS_ST, tbVendasDet.Valor_ICMS_ST, tbVendasDet.Aliq_IPI, tbVendasDet.Valor_IPI, tbVendasDet.Aliq_PIS, tbVendasDet.Aliq_Cofins, tbVendasDet.Valor_PIS, tbVendasDet.Valor_Cofins " & _
'"FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda " & _
'"WHERE dataemissao >= #" & cDtINI & "# and dataemissao <= #" & cDtFIM & "# ORDER BY tbVendas.ID, tbVendasDet.ID;")



Dim rsRegSaida As DAO.Recordset
Set rsRegSaida = Db.OpenRecordset("SELECT NatOperacao, Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbvendasdet.CFOP_ESCRITURADA AS CFOP, tbvendasdet.CFOP_ESC_DESC AS [CFOP Desc], tbvendasdet.lancfiscal AS [Lanc Fiscal], Sum(tbvendasdet.ValorTot) AS [Valor Contabil], Sum(tbvendasdet.BaseCalculo) AS [Base de Calculo], Sum(tbvendasdet.Valor_ICMS) AS ICMS, Sum(tbvendasdet.Valor_IPI) AS IPI, Sum(tbvendasdet.Valor_PIS) AS PIS, Sum(tbvendasdet.Valor_Cofins) AS Cofins, Sum(tbvendasdet.Valor_ICMS_ST) AS [ICMS ST], Sum(tbvendasdet.BaseCalc_ST) AS BaseCalc_ST, tbvendasdet.CST, tbvendasdet.CST_DESC, tbvendasdet.Aliq_ICMS " & _
"FROM (tbCliente INNER JOIN tbvendas ON tbCliente.IDCliente = tbvendas.Idcliente) INNER JOIN (tbCadProd INNER JOIN tbvendasdet ON (tbCadProd.IDProd = tbvendasdet.IDProd) AND (tbCadProd.IDProd = tbvendasdet.IDProd)) ON tbvendas.ID = tbvendasdet.IDVenda " & _
"WHERE tbVendas.DataEmissao >= #" & cDtIniVb & "# And tbVendas.DataEmissao <= #" & cDtFimVb & "# and [STATUS] = 'ATIVO' " & _
"GROUP BY Year(DataEmissao), Month(DataEmissao), tbvendasdet.CFOP_ESCRITURADA, tbvendasdet.CFOP_ESC_DESC, tbvendasdet.lancfiscal, tbvendas.TipoNF, tbvendasdet.CST, tbvendasdet.CST_DESC, tbvendasdet.Aliq_ICMS, NatOperacao " & _
"HAVING tbvendas.TipoNF = '1-SAIDA' and tbvendas.NatOperacao <> 'Venda Cupom Fiscal SAT' " & _
"ORDER BY Year(DataEmissao), Month(DataEmissao);")



Dim rsRegEntr As DAO.Recordset

Dim rsEnergia As DAO.Recordset
Set rsEnergia = Db.OpenRecordset("SELECT * FROM tbCompras WHERE (((tbCompras.IdFornecedor)=1131) AND ((tbCompras.DataEmissao)>= #" & cDtIniVb & "#  And (tbCompras.DataEmissao)<= #" & cDtFimVb & "#));")

Dim rsEnergiaDet As DAO.Recordset
Set rsEnergiaDet = Db.OpenRecordset("SELECT tbComprasDet.* FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra WHERE (((tbCompras.IdFornecedor)=1131) AND ((tbCompras.DataEmissao)>=#" & cDtIniVb & "# And (tbCompras.DataEmissao)<=#" & cDtFimVb & "#));")


Dim rsConsVenda As DAO.Recordset
Set rsConsVenda = Db.OpenRecordset("SELECT tbvendas.NatOperacao, tbVendasDet.IDProd, tbCadProd.NCM, Sum(tbVendasDet.ValorTot) AS SomaDeValorTot, tbCadProd.DescProd, tbCadProd.PROD_FINAL, tbCadProd.REVENDA " & _
"FROM tbCadProd INNER JOIN (tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda) ON tbCadProd.IDProd = tbVendasDet.IDProd " & _
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIniVb & "# And (tbVendas.DataEmissao) <= #" & cDtFimVb & "#)) Or (((tbVendas.DataEmissao) >= #" & cDtIniVb & "# And (tbVendas.DataEmissao) <= #" & cDtFimVb & "#)) " & _
"GROUP BY tbVendasDet.IDProd, tbCadProd.NCM, tbCadProd.DescProd, tbCadProd.PROD_FINAL, tbCadProd.REVENDA, tbvendas.NatOperacao " & _
"HAVING tbCadProd.PROD_FINAL='SIM' and tbvendas.NatOperacao <> 'Venda Cupom Fiscal SAT'  OR tbCadProd.REVENDA='SIM' and tbvendas.NatOperacao <> 'Venda Cupom Fiscal SAT'  ;")


cLinC = 0

Dim cDtC100 As String
Dim cDtC500 As String
cDtC100 = "NAO"
cDtC500 = "NAO"

If rsCompra.RecordCount > 0 Then
cDtC100 = "SIM"
Else: End If

If rsVenda.RecordCount > 0 Then
cDtC100 = "SIM"
Else: End If

If rsEnergia.RecordCount > 0 Then
cDtC500 = "SIM"
Else: End If



'cC001
If cDtC100 = "SIM" Or cDtC500 = "SIM" Then
cC001 = "|" & "C001" & "|" & "0" & "|"
cLinC = cLinC + 1
lC001 = lC001 + 1
Print #iArq, cC001
Else
cC001 = "|" & "C001" & "|" & "1" & "|"
cLinC = cLinC + 1
lC001 = lC001 + 1
Print #iArq, cC001
GoTo SemDadosC001
End If


'Registro C010: Identificação do Estabelecimento
'cC010
cC010 = "|" & "C010" & "|" & rsEmpresa!CNPJ & "|" & "2" & "|"
cLinC = cLinC + 1
lC010 = lC010 + 1
Print #iArq, cC010


'CREDITO DE COMPRA APENAS LUCRO REAL
'GoTo pulaC100Compra:
'Registro  C100:  Documento- Nota  Fiscal  (Código  01),  Nota  Fiscal Avulsa  (Código  1B),  Nota Fiscal de Produtor (Código 04), NF-e (Código 55) e NFC-e (Código 65).
'NOTAS DE ENTRADA COMPRA
'cC100
Dim clin170 As Integer
Dim cID170 As Integer

Do Until rsCompra.EOF
If rsCompra!IdFornecedor = 1131 Then
Else

'cC100 = "|" & "C100" & "|" & "0" & "|" & "1" & "|" & rsCompra!IdFornecedor & "|" & "55" & "|" & "00" & "|" & rsCompra!Serie & "|" & rsCompra!NumNF & "|" & rsCompra!ChaveNF & "|" & Replace(rsCompra!DataEmissao, "/", "") & "|" & Replace(rsCompra!DataEmissao, "/", "") & "|" & Round(rsCompra!VlrTOTALNF, 2) & "|" & "0" & "|" & Round(rsCompra!VlrDesconto, 2) & "||" & Round(rsCompra!VlrTotalProdutos, 2) & "|" & "9" & "|" & "|" & "|" & "|" & Round(rsCompra!ICMS_BaseCalc, 2) & "|" & Round(rsCompra!ICMS_Valor, 2) & "|" & Round(rsCompra!ICMS_ST_BaseCalc, 2) & "|" & Round(rsCompra!ICMS_ST_Valor, 2) & "|" & Round(rsCompra!IPI_Valor, 2) & "|" & Round(rsCompra!PIS_Valor, 2) & "|" & Round(rsCompra!Cofins_Valor, 2) & "|" & "|" & "|"
'Print #iArq, cC100
'cLinC = cLinC + 1
'lC100 = lC100 + 1

    'cC170
    'COMPRAS
    cC170 = ""
    cID170 = rsCompra!ID
    clin170 = 1
    
    Set rsCompraDet = Db.OpenRecordset("SELECT tbCompras.ID as ID, tbComprasDet.CST_ICMS, tbComprasDet.CST_IPI, tbComprasDet.CST_PIS, tbComprasDet.CST_Cofins, tbComprasDet.ID as ID_DET,tbComprasDet.IDCompra as ID_Compra,  tbComprasDet.IDProd, tbComprasDet.Qnt, tbCadProd.Unid, tbCompras.dataemissao, tbComprasDet.ValorTot, tbComprasDet.VlrDesc, tbComprasDet.CST, tbComprasDet.CFOP, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.BaseCalculo,tbComprasDet.BaseCalc_PisCofins, tbComprasDet.Aliq_ICMS, tbComprasDet.Valor_ICMS, tbComprasDet.BaseCalc_ST, tbComprasDet.Aliq_ICMS_ST, tbComprasDet.Valor_ICMS_ST, tbComprasDet.Aliq_IPI, tbComprasDet.Valor_IPI, tbComprasDet.Aliq_PIS, tbComprasDet.Aliq_Cofins, tbComprasDet.Valor_PIS, tbComprasDet.Valor_Cofins " & _
    "FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
    "WHERE dataemissao >= #" & cDtIniVb & "# and dataemissao <= #" & cDtFimVb & "# and tbComprasDet.IDCompra = " & cID170 & " and LancFiscal = 'CREDITO' ORDER BY tbCompras.ID, tbComprasDet.ID;")
    
    If rsCompraDet.EOF And rsCompraDet.BOF Then
    GoTo PulaRsCompraDet:
    Else
    cC100 = "|" & "C100" & "|" & "0" & "|" & "1" & "|" & rsCompra!IdFornecedor & "|" & "55" & "|" & "00" & "|" & rsCompra!Serie & "|" & rsCompra!NumNF & "|" & rsCompra!chavenf & "|" & Replace(rsCompra!DataEmissao, "/", "") & "|" & Replace(rsCompra!DataEmissao, "/", "") & "|" & Round(rsCompra!VlrTOTALNF, 2) & "|" & "0" & "|" & Round(rsCompra!VlrDesconto, 2) & "||" & Round(rsCompra!VlrTotalProdutos, 2) & "|" & "9" & "|" & "|" & "|" & "|" & Round(rsCompra!ICMS_BaseCalc, 2) & "|" & Round(rsCompra!ICMS_Valor, 2) & "|" & Round(rsCompra!ICMS_ST_BaseCalc, 2) & "|" & Round(rsCompra!ICMS_ST_Valor, 2) & "|" & Round(rsCompra!IPI_Valor, 2) & "|" & Round(rsCompra!PIS_Valor, 2) & "|" & Round(rsCompra!COFINS_Valor, 2) & "|" & "|" & "|"
    Print #iArq, cC100
    cLinC = cLinC + 1
    lC100 = lC100 + 1
    End If
    
    clin170 = 1
    Do Until rsCompraDet.EOF
    
    cC170 = "|" & "C170" & "|" & clin170 & "|" & rsCompraDet!IDProd & "||" & rsCompraDet!Qnt & "|" & rsCompraDet!Unid & "|" & Round(rsCompraDet!ValorTot, 2) & "|" & Round(rsCompraDet!VlrDesc, 2) & "|" & "0" & "|" & rsCompraDet!CST_ICMS & "|" & rsCompraDet!CFOP_ESCRITURADA & "|" & rsCompraDet!CFOP_ESCRITURADA & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_ICMS & "|" & Round(rsCompraDet!Valor_ICMS, 2) & "|" & Round(rsCompraDet!BaseCalc_ST, 2) & "|" & rsCompraDet!Aliq_ICMS_ST & "|" & Round(rsCompraDet!Valor_ICMS_ST, 2) & "|" & "0" & "|" & rsCompraDet!CST_IPI & "|" & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_IPI & "|" & Round(rsCompraDet!Valor_IPI, 2) & "|" & rsCompraDet!CST_PIS & "|" & Round(rsCompraDet!BaseCalc_PisCofins, 2) & "|" & rsCompraDet!Aliq_PIS & "|||" & Round(rsCompraDet!Valor_PIS, 2) & "|" & rsCompraDet!CST_Cofins & "|" & Round(rsCompraDet!BaseCalc_PisCofins, 2) & "|" & rsCompraDet!Aliq_Cofins & "|||" & Round(rsCompraDet!Valor_Cofins, 2) & "|" & "1" & "|"
    Print #iArq, cC170
    clin170 = clin170 + 1
    rsCompraDet.MoveNext
    cLinC = cLinC + 1
    lC170 = lC170 + 1
    Loop
    'cC170
    
    
    rsCompraDet.Close
    
PulaRsCompraDet:
    
End If
rsCompra.MoveNext
Loop

pulaC100Compra:

'NOTAS DE SAÍDA VENDA
'cC100
Do Until rsVenda.EOF
cC100 = "|" & "C100" & "|" & "1" & "|" & "0" & "|" & rsVenda!IdCliente & "|" & "55" & "|" & "00" & "|" & rsVenda!Serie & "|" & rsVenda!NumNF & "|" & rsVenda!chavenf & "|" & Left(Replace(rsVenda!DataEmissao, "/", ""), 8) & "|" & Left(Replace(rsVenda!DataEmissao, "/", ""), 8) & "|" & Round(rsVenda!VlrTOTALNF, 2) & "|" & "0" & "|" & Round(rsVenda!VlrDesconto, 2) & "||" & Round(rsVenda!VlrTotalProdutos - rsVenda!VlrDesconto, 2) & "|" & "9" & "|" & "|" & "|" & "|" & Round(rsVenda!ICMS_BaseCalc, 2) & "|" & Round(rsVenda!ICMS_Valor, 2) & "|" & Round(rsVenda!ICMS_ST_BaseCalc, 2) & "|" & Round(rsVenda!ICMS_ST_Valor, 2) & "|" & Round(rsVenda!IPI_Valor, 2) & "|" & Round(rsVenda!PIS_Valor, 2) & "|" & Round(rsVenda!COFINS_Valor, 2) & "|" & "|" & "|"
Print #iArq, cC100
cLinC = cLinC + 1
lC100 = lC100 + 1

'VENDAS
    'cC170
    cC170 = ""
    cID170 = rsVenda!ID
    clin170 = 1
    
    Set rsVendaDet = Db.OpenRecordset("SELECT tbVendas.ID as ID, tbVendasDet.ID as ID_DET,tbVendasDet.IDVenda as ID_Venda,  tbVendasDet.IDProd, tbVendasDet.Qnt, tbCadProd.Unid, tbVendas.dataemissao, tbVendasDet.ValorTot, tbVendasDet.VlrDesc, tbVendasDet.CST, tbVendasDet.CST_ICMS, tbVendasDet.CST_IPI, tbVendasDet.CST_PIS, tbVendasDet.CST_Cofins, tbVendasDet.CFOP, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.BaseCalculo,tbVendasDet.BaseCalc_PisCofins, tbVendasDet.Aliq_ICMS, tbVendasDet.Valor_ICMS, tbVendasDet.BaseCalc_ST, tbVendasDet.Aliq_ICMS_ST, tbVendasDet.Valor_ICMS_ST, tbVendasDet.Aliq_IPI, tbVendasDet.Valor_IPI, tbVendasDet.Aliq_PIS, tbVendasDet.Aliq_Cofins, tbVendasDet.Valor_PIS, tbVendasDet.Valor_Cofins " & _
    "FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda " & _
    "WHERE dataemissao >= #" & cDtIniVb & "# and dataemissao <= #" & cDtFimVb & " 23:59:59" & "# and tbVendasDet.IDVenda = " & cID170 & " and [Status] = 'ATIVO' ORDER BY tbVendas.ID, tbVendasDet.ID;")

    clin170 = 1
    Do Until rsVendaDet.EOF
    
    cC170 = "|" & "C170" & "|" & clin170 & "|" & rsVendaDet!IDProd & "||" & rsVendaDet!Qnt & "|" & rsVendaDet!Unid & "|" & Round(rsVendaDet!ValorTot - rsVendaDet!VlrDesc, 2) & "|" & Round(rsVendaDet!VlrDesc, 2) & "|" & "0" & "|" & rsVendaDet!CST_ICMS & "|" & rsVendaDet!CFOP_ESCRITURADA & "|" & rsVendaDet!CFOP_ESCRITURADA & "|" & Round(rsVendaDet!BaseCalculo, 2) & "|" & rsVendaDet!Aliq_ICMS & "|" & Round(rsVendaDet!Valor_ICMS, 2) & "|" & Round(rsVendaDet!BaseCalc_ST, 2) & "|" & rsVendaDet!Aliq_ICMS_ST & "|" & Round(rsVendaDet!Valor_ICMS_ST, 2) & "|" & "0" & "|" & rsVendaDet!CST_IPI & "|" & "|" & Round(rsVendaDet!BaseCalculo, 2) & "|" & rsVendaDet!Aliq_IPI & "|" & Round(rsVendaDet!Valor_IPI, 2) & "|" & rsVendaDet!CST_PIS & "|" & Round(rsVendaDet!BaseCalc_PisCofins, 2) & "|" & rsVendaDet!Aliq_PIS & "|||" & Round(rsVendaDet!Valor_PIS, 2) & "|" & rsVendaDet!CST_Cofins & "|" & Round(rsVendaDet!BaseCalc_PisCofins, 2) & "|" & rsVendaDet!Aliq_Cofins & "|||" & Round(rsVendaDet!Valor_Cofins, 2) & "|" & "1" & "|"
    'cC170 = "|" & "C170" & "|" & clin170 & "|" & rsCompraDet!IDProd & "||" & rsCompraDet!Qnt & "|" & rsCompraDet!Unid & "|" & Round(rsCompraDet!ValorTot, 2) & "|" & Round(rsCompraDet!VlrDesc, 2) & "|" & "0" & "|" & rsCompraDet!CST_ICMS & "|" & rsCompraDet!CFOP_ESCRITURADA & "|" & rsCompraDet!CFOP_ESCRITURADA & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_ICMS & "|" & Round(rsCompraDet!Valor_ICMS, 2) & "|" & Round(rsCompraDet!BaseCalc_ST, 2) & "|" & rsCompraDet!Aliq_ICMS_ST & "|" & Round(rsCompraDet!Valor_ICMS_ST, 2) & "|" & "0" & "|" & rsCompraDet!CST_IPI & "|" & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_IPI & "|" & Round(rsCompraDet!Valor_IPI, 2) & "|" & rsCompraDet!CST_PIS & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_PIS & "|||" & Round(rsCompraDet!Valor_PIS, 2) & "|" & rsCompraDet!CST_Cofins & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_COFINS & "|||" & Round(rsCompraDet!Valor_Cofins, 2) & "|" & "1" & "|"
    
    Print #iArq, cC170
    clin170 = clin170 + 1
    rsVendaDet.MoveNext
    cLinC = cLinC + 1
    lC170 = lC170 + 1
    Loop
    
    'c175 similar ao c190 da EFD ICMS
    
   

    rsVendaDet.Close

rsVenda.MoveNext

'lC100 = lC100 + 1
Loop


'Registro C180: Consolidação de Notas Fiscais Eletrônicas Emitidas Pela Pessoa Jurídica
'cC180
Dim rsDetPIS As Recordset
Dim rsDetCofins As Recordset

Do Until rsConsVenda.EOF

    cC180 = "|" & "C180" & "|" & "55" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|" & rsConsVenda!IDProd & "|" & rsConsVenda!NCM & "|" & "|" & Round(rsConsVenda!SomaDeValorTot, 2) & "|"
    cLinC = cLinC + 1
    lC180 = lC180 + 1
    Print #iArq, cC180
    
    'Registro C181: Detalhamento da Consolidação – Operações de Vendas – PIS/Pasep
    
    Set rsDetPIS = Db.OpenRecordset("SELECT tbVendasDet.CST_PIS, tbVendasDet.CFOP_ESCRITURADA AS CFOP, Sum(tbVendasDet.ValorTot) AS ValorTot, Sum(tbVendasDet.VlrDesc) AS VlrDesc, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, tbVendasDet.Aliq_PIS, Sum(tbVendasDet.Valor_PIS) AS Valor_PIS, tbCadProd.PROD_FINAL, tbCadProd.REVENDA " & _
    "FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
    "WHERE (tbVendasDet.IDProd = " & rsConsVenda!IDProd & " And tbVendas.DataEmissao >= #" & cDtIniVb & "# And tbVendas.DataEmissao <= #" & cDtFimVb & "# and [status] = 'ATIVO') Or (tbVendasDet.IDProd = " & rsConsVenda!IDProd & " And tbVendas.DataEmissao >= #" & cDtIniVb & "# And tbVendas.DataEmissao <= #" & cDtFimVb & "# AND [STATUS] = 'ATIVO') " & _
    "GROUP BY tbVendasDet.CST_PIS, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.Aliq_PIS, tbCadProd.PROD_FINAL, tbCadProd.REVENDA; ")

    cC181 = "|" & "C181" & "|" & rsDetPIS!CST_PIS & "|" & rsDetPIS!CFOP & "|" & Round(rsDetPIS!ValorTot, 2) & "|" & Round(rsDetPIS!VlrDesc, 2) & "|" & Round(rsDetPIS!ValorTot, 2) & "|" & rsDetPIS!Aliq_PIS & "|" & "|" & "|" & Round(rsDetPIS!Valor_PIS, 2) & "|" & "2" & "|"
    cLinC = cLinC + 1
    lC181 = lC181 + 1
    Print #iArq, cC181
    
    rsDetPIS.Close
    'Registro C185: Detalhamento da Consolidação–Operações de Vendas–Cofins
    
    Set rsDetCofins = Db.OpenRecordset("SELECT tbVendasDet.CST_Cofins, tbVendasDet.CFOP_ESCRITURADA AS CFOP, Sum(tbVendasDet.ValorTot) AS ValorTot, Sum(tbVendasDet.VlrDesc) AS VlrDesc, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, tbVendasDet.Aliq_Cofins, Sum(tbVendasDet.Valor_Cofins) AS Valor_Cofins " & _
    "FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
    "WHERE (tbVendasDet.IDProd = " & rsConsVenda!IDProd & " And tbVendas.DataEmissao >= # " & cDtIniVb & " # And tbVendas.DataEmissao <= #" & cDtFimVb & "#) Or (tbVendasDet.IDProd = " & rsConsVenda!IDProd & " And tbVendas.DataEmissao >= #" & cDtIniVb & "# And tbVendas.DataEmissao <= #" & cDtFimVb & "# and [STATUS] = 'ATIVO') " & _
    "GROUP BY tbVendasDet.CST_Cofins, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.Aliq_Cofins;")
    
    cC185 = "|" & "C185" & "|" & rsDetCofins!CST_Cofins & "|" & rsDetCofins!CFOP & "|" & Round(rsDetCofins!ValorTot, 2) & "|" & Round(rsDetCofins!VlrDesc, 2) & "|" & Round(rsDetCofins!ValorTot, 2) & "|" & rsDetCofins!Aliq_Cofins & "|" & "|" & "|" & Round(rsDetCofins!Valor_Cofins, 2) & "|" & "2" & "|"
    cLinC = cLinC + 1
    lC185 = lC185 + 1
    Print #iArq, cC185
    
    rsDetCofins.Close
        
rsConsVenda.MoveNext
Loop


'Registro C500: Nota Fiscal/Conta de Energia Elétrica (Código 06),
'cC500
'C500
cC500 = ""
Do Until rsEnergia.EOF

cC500 = "|" & "C500" & "|" & rsEnergia!IdFornecedor & "|" & "06" & "|" & "00" & "|" & rsEnergia!Serie & "|" & "|" & rsEnergia!NumNF & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & rsEnergia!VlrTotalProdutos & "|" & rsEnergia!ICMS_Valor & "|" & "|" & rsEnergia!PIS_Valor & "|" & rsEnergia!COFINS_Valor & "|" & "|"
Print #iArq, cC500
cLinC = cLinC + 1
lC500 = lC500 + 1

'cC501
Set rsEnerdet501 = Db.OpenRecordset("SELECT tbComprasDet.CST_PIS, Sum(tbComprasDet.ValorTot) AS ValorTot, Sum(tbComprasDet.BaseCalculo) AS BaseCalculo, tbComprasDet.Aliq_PIS, Sum(tbComprasDet.Valor_PIS) AS Valor_PIS, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"GROUP BY tbComprasDet.CST_PIS, tbComprasDet.Aliq_PIS, tbCompras.ChaveNF " & _
"HAVING (((tbCompras.ChaveNF)='" & rsEnergia!chavenf & "'));")

cC501 = "|" & "C501" & "|" & rsEnerdet501!CST_PIS & "|" & rsEnerdet501!ValorTot & "|" & "04" & "|" & rsEnerdet501!BaseCalculo & "|" & rsEnerdet501!Aliq_PIS & "|" & Round(rsEnerdet501!Valor_PIS, 2) & "|" & "2" & "|"
Print #iArq, cC501
cLinC = cLinC + 1
lC501 = lC501 + 1

'cC505
Set rsEnerdet505 = Db.OpenRecordset("SELECT tbComprasDet.CST_COFINS, Sum(tbComprasDet.ValorTot) AS ValorTot, Sum(tbComprasDet.BaseCalculo) AS BaseCalculo, tbComprasDet.Aliq_Cofins, Sum(tbComprasDet.Valor_Cofins) AS Valor_Cofins, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"GROUP BY tbComprasDet.CST_COFINS, tbComprasDet.Aliq_Cofins, tbCompras.ChaveNF " & _
"HAVING (((tbCompras.ChaveNF)='" & rsEnergia!chavenf & "'));")

cC505 = "|" & "C505" & "|" & rsEnerdet505!CST_Cofins & "|" & rsEnerdet505!ValorTot & "|" & "04" & "|" & rsEnerdet505!BaseCalculo & "|" & "7,6" & "|" & Round(rsEnerdet505!Valor_Cofins, 2) & "|" & "2" & "|"
Print #iArq, cC505
cLinC = cLinC + 1
lC505 = lC505 + 1

rsEnerdet501.Close
rsEnerdet505.Close



rsEnergia.MoveNext
Loop

'Registro C501: Complemento da Operação (Códigos 06, 28 e 29)–PIS/Pasep
'cC501
'C501



'C510
'cC510 = ""
'cLin510 = 1
'Do Until rsEnergiaDet.EOF
'cC510 = "|" & "C510" & "|" & cLin510 & "|" & rsEnergiaDet!IdProd & "|" & "0601" & "|" & rsEnergiaDet!Qnt & "|" & "KW" & "|" & rsEnergiaDet!ValorTot & "|" & rsEnergiaDet!VlrDesc & "|" & rsEnergiaDet!CST_ICMS & "|" & rsEnergiaDet!CFOP_Escriturada & "|" & rsEnergiaDet!BaseCalculo & "|" & rsEnergiaDet!Aliq_ICMS & "|" & rsEnergiaDet!Valor_ICMS & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "1" & "|" & "1131" & "|" & rsEnergiaDet!Valor_PIS & "|" & rsEnergiaDet!Valor_Cofins & "|" & "2" & "|"
'Print #iArq, cC510
'cLin510 = cLin510 + 1
'lC510 = lC510 + 1
'cLinC = cLinC + 1
'rsEnergiaDet.MoveNext
'Loop

If rsEnergiaDet.RecordCount > 0 Then
rsEnergiaDet.MoveFirst
Else: End If

Do Until rsEnergiaDet.EOF
'C590 Consolidação de Nf de energia
cC590 = ""
cC590 = "|" & "C590" & "|" & rsEnergiaDet!CST_ICMS & "|" & rsEnergiaDet!CFOP_ESCRITURADA & "|" & rsEnergiaDet!Aliq_ICMS & "|" & rsEnergiaDet!ValorTot & "|" & rsEnergiaDet!BaseCalculo & "|" & rsEnergiaDet!Valor_ICMS & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "|"
'Print #iArq, cC590
'Na EFD Contribuições não pede o registro C590
rsEnergiaDet.MoveNext
'cLinC = cLinC + 1
'lC590 = lC590 + 1
Loop

'Registro C860: Identificação do Equipamento SAT-CF-e
cC860 = ""
Set rsSAT = Db.OpenRecordset("select format(tbVendasSAT.DataEmissao,'dd/mm/yyyy') as DataEmissao, NumSerieSAT, min(numCF) as MinCF, max(numCF) as MaxCF from tbVendasSAT WHERE tbVendasSAT.DataEmissao  >= # " & cDtIni & " 00:00:00" & " # And tbVendasSAT.DataEmissao <= # " & cDtFim & " 23:59:59" & " # GROUP BY format(DataEmissao,'dd/mm/yyyy'), NumSerieSAT;")
Do Until rsSAT.EOF
cC860 = "|" & "C860" & "|" & "59" & "|" & rsSAT!NumSerieSAT & "|" & Replace(Format(rsSAT!DataEmissao, "dd/mm/yyyy"), "/", "") & "|" & rsSAT!MinCF & "|" & rsSAT!MaxCF & "|"
Print #iArq, cC860
cLinC = cLinC + 1
lC860 = lC860 + 1

    'C870
    cC870 = ""
    Set rsSATDet = Db.OpenRecordset("select q2.CFOP, q2.CST_PIS, q2.CST_COFINS, q2.pPIS, q2.pCofins, q2.IdProd, " & _
                                    "sum(q2.Vlr_Item) as Vlr_Item, sum(q2.Vlr_Desc) as Vlr_Desc, sum(q2.bCalcPIS) as bCalcPIS, " & _
                                    "sum(q2.vPIS) as vPIS, sum(q2.bCalcCofins) as bCalcCofins, sum(q2.vCofins) as vCofins  from " & _
                                    "(select format(DataEmissao,'dd/mm/yyyy') as DataEmissaos, IDSat from tbVendasSAT " & _
                                    "where format(DataEmissao,'dd/mm/yyyy') = '" & Format(rsSAT!DataEmissao, "dd/mm/yyyy") & "' " & _
                                    "group by IDSat, format(DataEmissao,'dd/mm/yyyy')) as q1 " & _
                                    "inner Join " & _
                                    "(select * from tbVendasSATDet) as q2 " & _
                                    "on q1.IDSat = q2.IDSat " & _
                                    "group by q2.CFOP, q2.CST_PIS, q2.CST_COFINS, q2.pPIS, q2.pCofins, q2.IdProd;")


    Do Until rsSATDet.EOF
    cC870 = "|" & "C870" & "|" & rsSATDet!IDProd & "|" & rsSATDet!CFOP & "|" & rsSATDet!Vlr_Item & "|" & rsSATDet!Vlr_Desc & "|" & rsSATDet!CST_PIS & "|" & rsSATDet!bCalcPIS & "|" & rsSATDet!pPIS & "|" & rsSATDet!vPIS & "|" & rsSATDet!CST_Cofins & "|" & rsSATDet!bCalcCofins & "|" & rsSATDet!pCofins & "|" & rsSATDet!vCofins & "|" & "4" & "|"
    Print #iArq, cC870
    rsSATDet.MoveNext
    cLinC = cLinC + 1
    lC870 = lC870 + 1
    Loop
    
           
rsSAT.MoveNext
Loop


SemDadosC001:


'C990
cLinC = cLinC + 1
lC990 = lC990 + 1
cC990 = "|" & "C990" & "|" & cLinC & "|"
Print #iArq, cC990


'BLOCO D: Documentos Fiscais–II-Serviços (ICMS)
Set rsTransp = Db.OpenRecordset("SELECT tbTransportes.*, tbMunicipioIBGE.COD_MUNICIPIO AS Rem_IBGE, tbMunicipioIBGE_1.COD_MUNICIPIO AS Dest_IBGE " & _
"FROM (tbTransportes INNER JOIN tbMunicipioIBGE ON tbTransportes.RemetenteCidade = tbMunicipioIBGE.MUNICIPIO_ACENTO) INNER JOIN tbMunicipioIBGE AS tbMunicipioIBGE_1 ON tbTransportes.DestinatarioCidade = tbMunicipioIBGE_1.MUNICIPIO_ACENTO " & _
"WHERE (((tbTransportes.DataEmissao)>=#" & cDtIniVb & "# And (tbTransportes.DataEmissao)<=#" & cDtFimVb & "#));")

Dim rsTrTotal As DAO.Recordset
Dim cLinBlocoD As Integer
cLinBlocoD = 0




'cD001
flagTrp = 0
If rsTransp.EOF Then
cD001 = "|" & "D001" & "|" & "1" & "|"
cLinBlocoD = cLinBlocoD + 1
lD001 = lD001 + 1
Print #iArq, cD001
flagTrp = 1
GoTo semdatatransportes
Else
'CREDITO APENAS PARA LUCRO REAL
cD001 = "|" & "D001" & "|" & "1" & "|"
cLinBlocoD = cLinBlocoD + 1
lD001 = lD001 + 1
Print #iArq, cD001
flagTrp = 1
GoTo semdatatransportes
End If


'cD010
cD010 = "|" & "D010" & "|" & "23866944000141" & "|"
cLinBlocoD = cLinBlocoD + 1
lD010 = lD010 + 1
Print #iArq, cD010


'cD100

Do Until rsTransp.EOF
   ' If cCodVer = "011" Then
    'cD100 = "|" & "D100" & "|" & "0" & "|" & "1" & "|" & rsTransp!ID_Emit & "|" & "57" & "|" & "00" & "|" & rsTransp!Serie & "|" & "|" & rsTransp!Num_CTE & "|" & rsTransp!ChaveCTE & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & "|" & rsTransp!ChaveCTE & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & "0" & "|" & "1" & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & Round(rsTransp!BaseCalcICMS, 2) & "|" & Round(rsTransp!ValorICMS, 2) & "|" & "0" & "|" & "|" & "2" & "|"
    'ElseIf cCodVer = "012" Then
    
    cD100 = "|" & "D100" & "|" & "0" & "|" & "1" & "|" & rsTransp!ID_Emit & "|" & "57" & "|" & "00" & "|" & rsTransp!Serie & "|" & "|" & rsTransp!Num_CTE & "|" & rsTransp!ChaveCTE & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & "|" & rsTransp!ChaveCTE & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & "0" & "|" & "1" & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & Round(rsTransp!BaseCalcICMS, 2) & "|" & Round(rsTransp!ValorICMS, 2) & "|" & "0" & "|" & "|" & "2" & "|"
    'End If
    Print #iArq, cD100
    cLinBlocoD = cLinBlocoD + 1
    lD100 = lD100 + 1
    
    'cD101
    'Registro D101: Complemento do Documento de Transporte (Códigos 07, 08, 8B, 09, 10, 11, 26, 27 e
    'COMPRAS
    Set rsTr101_Compras = Db.OpenRecordset("SELECT tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbComprasDet.CST_PIS, tbComprasDet.CFOP, Sum(tbComprasDet.ValorTot) AS ValorTot, Sum(tbComprasDet.BaseCalculo) AS BaseCalculo, tbComprasDet.Aliq_PIS, Sum(tbComprasDet.Valor_PIS) AS Valor_PIS, Sum(tbComprasDet.Qnt) AS Qnt, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "FROM tbTransportes INNER JOIN ((tbTransportesDet INNER JOIN tbCompras ON tbTransportesDet.ChaveNFe = tbCompras.ChaveNF) INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbTransportes.ID = tbTransportesDet.ID_Transporte " & _
    "GROUP BY tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbComprasDet.CST_PIS, tbComprasDet.CFOP, tbComprasDet.Aliq_PIS, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "HAVING (((tbTransportes.ChaveCTE)='" & rsTransp!ChaveCTE & "'));")

            
    cIndNat = "9"
    If rsTr101_Compras!CONSUMO = "SIM" Then
    cIndNat = "3"
    cCodBC = "02"
    Else: End If
    If rsTr101_Compras!IMOBILIZADO = "SIM" Then
    cIndNat = "2"
    cCodBC = "09"
    Else: End If
    If rsTr101_Compras!EMBALAGEM = "SIM" Then
    cIndNat = "2"
    cCodBC = "02"
    Else: End If
    If rsTr101_Compras!revenda = "SIM" Then
    cIndNat = "2"
    cCodBC = "01"
    Else: End If
    If rsTr101_Compras!MAT_PRIMA = "SIM" Then
    cIndNat = "2"
    cCodBC = "02"
    Else: End If
    
    cCodBC = ""
        
    cD101 = "|" & "D101" & "|" & cIndNat & "|" & rsTr101_Compras!ValorTot & "|" & rsTr101_Compras!CST_PIS & "|" & cCodBC & "|" & rsTr101_Compras!BaseCalculo & "|" & rsTr101_Compras!Aliq_PIS & "|" & rsTr101_Compras!Valor_PIS & "|" & "2" & "|"
    Print #iArq, cD101
    cLinBlocoD = cLinBlocoD + 1
    lD101 = lD101 + 1
    rsTr101_Compras.Close
    
    'VENDAS
    
    Set rsTr101_Vendas = Db.OpenRecordset("SELECT tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbVendasDet.CST_PIS, tbVendasDet.CFOP, Sum(tbVendasDet.ValorTot) AS ValorTot, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, tbVendasDet.Aliq_PIS, Sum(tbVendasDet.Valor_PIS) AS Valor_PIS, Sum(tbVendasDet.Qnt) AS Qnt, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "FROM (tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) INNER JOIN (tbTransportes INNER JOIN tbTransportesDet ON tbTransportes.ID = tbTransportesDet.ID_Transporte) ON tbVendas.ChaveNF = tbTransportesDet.ChaveNFe " & _
    "GROUP BY tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbVendasDet.CST_PIS, tbVendasDet.CFOP, tbVendasDet.Aliq_PIS, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "HAVING (((tbTransportes.ChaveCTE)='" & rsTransp!ChaveCTE & "'));")

    If rsTr101_Vendas.RecordCount > 0 Then
    Else: GoTo rsTr101_Vendas_Vazio:
    End If
        
    cIndNat = "0"
    If rsTr101_Vendas!CONSUMO = "SIM" Then
    cIndNat = "0"
    cCodBC = "02"
    Else: End If
    If rsTr101_Vendas!IMOBILIZADO = "SIM" Then
    cIndNat = "0"
    cCodBC = "09"
    Else: End If
    If rsTr101_Vendas!EMBALAGEM = "SIM" Then
    cIndNat = "0"
    cCodBC = "02"
    Else: End If
    If rsTr101_Vendas!revenda = "SIM" Then
    cIndNat = "0"
    cCodBC = "01"
    Else: End If
    If rsTr101_Vendas!MAT_PRIMA = "SIM" Then
    cIndNat = "0"
    cCodBC = "02"
    Else: End If
    
    cCodBC = ""
    
    cD101 = "|" & "D101" & "|" & cIndNat & "|" & rsTr101_Vendas!ValorTot & "|" & rsTr101_Vendas!CST_PIS & "|" & cCodBC & "|" & rsTr101_Vendas!BaseCalculo & "|" & rsTr101_Vendas!Aliq_PIS & "|" & rsTr101_Vendas!Valor_PIS & "|" & "2" & "|"
    Print #iArq, cD101
    cLinBlocoD = cLinBlocoD + 1
    lD101 = lD101 + 1
    
rsTr101_Vendas_Vazio:
'    rsTr101_Vendas.Close
    
    
    'Registro D105: Complemento do Documento de Transporte - Cofins
     'COMPRAS
    Set rsTr105_Compras = Db.OpenRecordset("SELECT tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbComprasDet.CST_Cofins, tbComprasDet.CFOP, Sum(tbComprasDet.ValorTot) AS ValorTot, Sum(tbComprasDet.BaseCalculo) AS BaseCalculo, tbComprasDet.Aliq_Cofins, Sum(tbComprasDet.Valor_Cofins) AS Valor_Cofins, Sum(tbComprasDet.Qnt) AS Qnt, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "FROM tbTransportes INNER JOIN ((tbTransportesDet INNER JOIN tbCompras ON tbTransportesDet.ChaveNFe = tbCompras.ChaveNF) INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbTransportes.ID = tbTransportesDet.ID_Transporte " & _
    "GROUP BY tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbComprasDet.CST_Cofins, tbComprasDet.CFOP, tbComprasDet.Aliq_Cofins, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "HAVING (((tbTransportes.ChaveCTE)='" & rsTransp!ChaveCTE & "'));")

            
    cIndNat = "9"
    If rsTr105_Compras!CONSUMO = "SIM" Then
    cIndNat = "3"
    cCodBC = "02"
    Else: End If
    If rsTr105_Compras!IMOBILIZADO = "SIM" Then
    cIndNat = "2"
    cCodBC = "09"
    Else: End If
    If rsTr105_Compras!EMBALAGEM = "SIM" Then
    cIndNat = "2"
    cCodBC = "02"
    Else: End If
    If rsTr105_Compras!revenda = "SIM" Then
    cIndNat = "2"
    cCodBC = "01"
    Else: End If
    If rsTr105_Compras!MAT_PRIMA = "SIM" Then
    cIndNat = "2"
    cCodBC = "02"
    Else: End If
    
    cCodBC = ""
    
    cD105 = "|" & "D105" & "|" & cIndNat & "|" & rsTr105_Compras!ValorTot & "|" & rsTr105_Compras!CST_Cofins & "|" & cCodBC & "|" & rsTr105_Compras!BaseCalculo & "|" & rsTr105_Compras!Aliq_Cofins & "|" & rsTr105_Compras!Valor_Cofins & "|" & "2" & "|"
    Print #iArq, cD105
    cLinBlocoD = cLinBlocoD + 1
    lD105 = lD105 + 1
    rsTr105_Compras.Close
    
    

    'VENDAS
     
    Set rsTr105_Vendas = Db.OpenRecordset("SELECT tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbVendasDet.CST_Cofins, tbVendasDet.CFOP, Sum(tbVendasDet.ValorTot) AS ValorTot, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, tbVendasDet.Aliq_Cofins, Sum(tbVendasDet.Valor_Cofins) AS Valor_Cofins, Sum(tbVendasDet.Qnt) AS Qnt, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "FROM (tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) INNER JOIN (tbTransportes INNER JOIN tbTransportesDet ON tbTransportes.ID = tbTransportesDet.ID_Transporte) ON tbVendas.ChaveNF = tbTransportesDet.ChaveNFe " & _
    "GROUP BY tbTransportes.ChaveCTE, tbTransportesDet.ChaveNFe, tbVendasDet.CST_Cofins, tbVendasDet.CFOP, tbVendasDet.Aliq_Cofins, tbCadProd.CONSUMO, tbCadProd.IMOBILIZADO, tbCadProd.PROD_FINAL, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tbCadProd.REVENDA " & _
    "HAVING (((tbTransportes.ChaveCTE)='" & rsTransp!ChaveCTE & "'));")

    If rsTr105_Vendas.RecordCount > 0 Then
    Else: GoTo rsTr105_Vendas_Vazio:
    End If
    cIndNat = "0"
    If rsTr101_Vendas!CONSUMO = "SIM" Then
    cIndNat = "0"
    cCodBC = "02"
    Else: End If
    If rsTr101_Vendas!IMOBILIZADO = "SIM" Then
    cIndNat = "0"
    cCodBC = "09"
    Else: End If
    If rsTr101_Vendas!EMBALAGEM = "SIM" Then
    cIndNat = "0"
    cCodBC = "02"
    Else: End If
    If rsTr101_Vendas!revenda = "SIM" Then
    cIndNat = "0"
    cCodBC = "01"
    Else: End If
    If rsTr101_Vendas!MAT_PRIMA = "SIM" Then
    cIndNat = "0"
    cCodBC = "02"
    Else: End If
    
    cCodBC = ""
    
    
    cD105 = "|" & "D105" & "|" & cIndNat & "|" & rsTr105_Vendas!ValorTot & "|" & rsTr105_Vendas!CST_Cofins & "|" & cCodBC & "|" & rsTr105_Vendas!BaseCalculo & "|" & rsTr105_Vendas!Aliq_Cofins & "|" & rsTr101_Vendas!Valor_Cofins & "|" & "2" & "|"
    Print #iArq, cD105
    cLinBlocoD = cLinBlocoD + 1
    lD105 = lD105 + 1
    
rsTr105_Vendas_Vazio:
'    rsTr105_Vendas.Close
     


rsTransp.MoveNext
'cLinBlocoD = cLinBlocoD + 1
'lD100 = lD100 + 1
Loop


semdatatransportes:

'cD990
cLinBlocoD = cLinBlocoD + 1
lD990 = lD990 + 1
cD990 = "|" & "D990" & "|" & cLinBlocoD & "|"
Print #iArq, cD990


'BLOCO F:Demais Documentos e Operações
'Registro F001: Abertura do Bloco F

'cF001
cF001 = "|" & "F001" & "|" & "0" & "|"
Print #iArq, cF001
cLinBlocoF = cLinBlocoF + 1
lF001 = lF001 + 1

'cF010
cF010 = "|" & "F010" & "|" & "23866944000141" & "|"
Print #iArq, cF010
cLinBlocoF = cLinBlocoF + 1
lF010 = lF010 + 1

'Registro  F120:  Bens  Incorporados  ao Ativo  Imobilizado – Operações  Geradoras  de  Créditos  com Base nos Encargos de Depreciação e Amortização
'Dim rsImob As DAO.Recordset
'Set rsImob = db.OpenRecordset("SELECT tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
'"FROM tbPlanoContasContabeis INNER JOIN (tbImobilizadoCadastro INNER JOIN tbCadProd_Ativo_temp ON tbImobilizadoCadastro.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta " & _
'"GROUP BY tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza;")

'Do Until rsImob.EOF
'cF120 = "|" & "F120" & "|" & "11" & "|" & "05" & "|" & "0" & "|" & "1" & "|" &
'rsImob.MoveNext

'Registro  F550:   Consolidação   das   Operações   da   Pessoa   Jurídica
'Submetida   ao Regime de Tributação com Base no Lucro  Presumido –
'Incidência do PIS/Pasep e da Cofins pelo Regime de Competência
'cF550
Dim rsF550 As Recordset
Set rsF550 = Db.OpenRecordset("SELECT tbVendas.DataEmissao, tbVendasDet.CST_PIS, tbVendasDet.CST_Cofins, tbVendasDet.CFOP, Sum(tbVendasDet.ValorTot) AS ValorTot, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, tbVendasDet.Aliq_PIS, Sum(tbVendasDet.Valor_PIS) AS Valor_PIS, tbVendasDet.Aliq_Cofins, Sum(tbVendasDet.Valor_Cofins) AS Valor_Cofins " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"GROUP BY tbVendas.DataEmissao, tbVendasDet.CST_PIS, tbVendasDet.CST_Cofins, tbVendasDet.CFOP, tbVendasDet.Aliq_PIS, tbVendasDet.Aliq_Cofins " & _
"HAVING (((tbVendas.DataEmissao)>=#" & cDtIniVb & "# And (tbVendas.DataEmissao)<=#" & cDtFimVb & " 23:59:59" & "#));")

Do Until rsF550.EOF

cF550 = "|" & "F550" & "|" & rsF550!ValorTot & "|" & rsF550!CST_PIS & "|" & "0" & "|" & rsF550!BaseCalculo & "|" & rsF550!Aliq_PIS & "|" & rsF550!Valor_PIS & "|" & rsF550!CST_Cofins & "|" & "0" & "|" & rsF550!BaseCalculo & "|" & rsF550!Aliq_Cofins & "|" & rsF550!Valor_Cofins & "|" & "55" & "|" & rsF550!CFOP & "|" & "2" & "|" & "|"
'Print #iArq, cF550
'cLinBlocoF = cLinBlocoF + 1
'lF550 = lF550 + 1
'NAO SEI PORQUE MAS DIZ NO SPED QUE NÃO É PRA FAZER ESSE REGISTRO
rsF550.MoveNext
Loop


'Registro  F560:   Consolidação   das   Operações   da   Pessoa   Jurídica
'Submetida   ao   Regime   de Tributação  com  Base  no  Lucro  Presumido –
'Incidência  do  PIS/Pasep  e  da  Cofins  pelo  Regime  de Competência
'(Apuração da Contribuição por Unidade de Medida de Produto – Alíquota em Reais)
'ESSE É PRA BEBIDAS FRIAS, COM VARIAÇÃO DE QUANTIDADE para fatos geradores até 30.04.2015
'para fatos geradores até 30.04.2015


'Registro  F990: Encerramento do Bloco F
'cF990
cLinBlocoF = cLinBlocoF + 1
lF990 = lF990 + 1
cF990 = "|" & "F990" & "|" & cLinBlocoF & "|"
Print #iArq, cF990

'BLOCO M–Apuração da Contribuição e Crédito do PIS/Pasep e da Cofins
'Registro M001: Abertura do Bloco M
'cM001


'Registro M210: Detalhamento da Contribuição para o PIS/Pasep do Período
Set rsM210 = Db.OpenRecordset("SELECT Round(Sum(Round([tbVendasDet].[ValorTot],2)-Round([tbVendasDet].[VlrDesc],2)+Round([tbVendasDet].[VlrOutro],2)),2) AS ValorTot, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, tbVendasDet.Aliq_PIS, Sum(tbVendasDet.Valor_PIS) AS Valor_PIS, tbVendasDet.CST_PIS " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIniVb & "# And (tbVendas.DataEmissao) <= #" & cDtFimVb & " 23:59:59" & "# and [STATUS]='ATIVO' AND TIPONF = '1-SAIDA')) " & _
"GROUP BY tbVendasDet.Aliq_PIS, tbVendasDet.CST_PIS;")

'Registro M610: Detalhamento da Contribuição para a Seguridade Social-Cofins do Período
Set rsM610 = Db.OpenRecordset("SELECT Round(Sum(Round([tbVendasDet].[ValorTot],2)-Round([tbVendasDet].[VlrDesc],2)+Round([tbVendasDet].[VlrOutro],2)),2) AS ValorTot, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, tbVendasDet.Aliq_Cofins, Sum(tbVendasDet.Valor_Cofins) AS Valor_Cofins, tbVendasDet.CST_Cofins " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIniVb & "# And (tbVendas.DataEmissao) <= #" & cDtFimVb & " 23:59:59" & "# AND [STATUS]='ATIVO' AND TIPONF = '1-SAIDA')) " & _
"GROUP BY tbVendasDet.ALIQ_COFINS, tbVendasDet.CST_COFINS;")

'0-Bloco com dados informados;
'1-Bloco sem dados informado

If rsVenda.RecordCount > 0 Then
cM001 = "|" & "M001" & "|" & "0" & "|"
cLinM = cLinM + 1
lM001 = lM001 + 1
Print #iArq, cM001
Else
cM001 = "|" & "M001" & "|" & "0" & "|"
cLinM = cLinM + 1
lM001 = lM001 + 1
Print #iArq, cM001
    
        'Registro M200: Consolidação da Contribuição para o PIS/Pasep do Período
        'cM200
        cM200 = "|" & "M200" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
        cLinM = cLinM + 1
        lM200 = lM200 + 1
        Print #iArq, cM200
        
        'Registro M600:  Consolidação da Contribuição para a Seguridade Social-Cofins do Período
        cM600 = "|" & "M600" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
        cLinM = cLinM + 1
        lM600 = lM600 + 1
        Print #iArq, cM600


GoTo semdadosM
End If


'Registro M105: Detalhamento da Base de Calculo do Crédito Apurado no Período - PISPasep
'APENAS PARA LUCRO REAL

'Registro M100: Crédito de PIS/Pasep Relativo ao Período
'APENAS PARA LUCRO REAL
'cM100

Set rsM100 = Db.OpenRecordset("SELECT  Sum(tbComprasDet.ValorTot-tbComprasDet.VlrDesc) AS ValorTot, Sum(tbComprasDet.BaseCalc_PisCofins) AS BaseCalc_PisCofins, tbComprasDet.Aliq_PIS, Sum(tbComprasDet.Valor_PIS) AS Valor_PIS, tbComprasDet.CST_PIS " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE tbCompras.DataEmissao >= #" & cDtIniVb & "# And tbCompras.DataEmissao <= #" & cDtFimVb & " 23:59:59" & "# and LancFiscal = 'CREDITO' " & _
"GROUP BY tbComprasDet.Aliq_PIS, tbComprasDet.CST_PIS;")

Dim vAliq_PIS As Double
If rsM100.EOF = True And rsM100.BOF = True Then
Else

Do Until rsM100.EOF

            vAliq_PIS = rsM100!Aliq_PIS
            
            If rsM100!Aliq_PIS = 1.65 Or rsM100!Aliq_PIS = 0.65 Then
            tAliq = 101
            Else
            tAliq = 102
            End If
            
            If rsM100!Aliq_PIS = 0 And rsM100!Valor_PIS > 0 Then
            vAliq_PIS = rsM100!Valor_PIS / rsM100!ValorTot
            Else
            End If
            
            
            
            
            cM100 = "|" & "M100" & "|" & tAliq & "|" & "0" & "|" & rsM100!BaseCalc_PisCofins & "|" & rsM100!Aliq_PIS & "|" & "|" & "|" & Round((rsM100!BaseCalc_PisCofins * rsM100!Aliq_PIS) / 100, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & Round((rsM100!BaseCalc_PisCofins * rsM100!Aliq_PIS) / 100, 2) & "|" & "0" & "|" & Round((rsM100!BaseCalc_PisCofins * rsM100!Aliq_PIS) / 100, 2) & "|" & "0" & "|"
            cLinM = cLinM + 1
            lM100 = lM100 + 1
            Print #iArq, cM100
            

Set rsM105 = Db.OpenRecordset("SELECT IIf(IdFornecedor=1131,'04',IIf(MAT_PRIMA='SIM','02',IIf(REVENDA='SIM','01',IIf(EMBALAGEM='SIM','02')))) AS Natureza, Sum(tbComprasDet.ValorTot-tbComprasDet.VlrDesc) AS ValorTot, Sum(tbComprasDet.BaseCalc_PisCofins) AS BaseCalc_PisCofins, tbComprasDet.Aliq_PIS, Sum(tbComprasDet.Valor_PIS) AS Valor_PIS, tbComprasDet.CST_PIS " & _
"FROM (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) INNER JOIN tbCadProd ON tbComprasDet.IDProd = tbCadProd.IDProd " & _
"WHERE tbCompras.DataEmissao>=#" & cDtIniVb & "# And tbCompras.DataEmissao<=#" & cDtFimVb & " 23:59:59" & "# AND tbComprasDet.LancFiscal='CREDITO' AND tbComprasDet.Aliq_PIS= " & Replace(rsM100!Aliq_PIS, ",", ".") & " " & _
"GROUP BY IIf(IdFornecedor=1131,'04',IIf(MAT_PRIMA='SIM','02',IIf(REVENDA='SIM','01',IIf(EMBALAGEM='SIM','02')))), tbComprasDet.Aliq_PIS, tbComprasDet.CST_PIS;")


If rsM105.EOF = True And rsM105.BOF = True Then
Else

Do Until rsM105.EOF

           
            
            cM105 = "|" & "M105" & "|" & rsM105!NATUREZA & "|" & rsM105!CST_PIS & "|" & Round(rsM105!BaseCalc_PisCofins, 2) & "|" & "0" & "|" & Round(rsM105!BaseCalc_PisCofins, 2) & "|" & Round(rsM105!BaseCalc_PisCofins, 2) & "|" & "|" & "|" & "Credito PIS" & "|"
            cLinM = cLinM + 1
            lM105 = lM105 + 1
            Print #iArq, cM105

rsM105.MoveNext
Loop
End If

rsM100.MoveNext
Loop
End If


Dim cVal210 As Double

If rsM210.RecordCount > 0 Then
'Registro M200: Consolidação da Contribuição para o PIS/Pasep do Período
'cM200

Do Until rsM210.EOF = True
cVal210 = cVal210 + rsM210!Valor_PIS
rsM210.MoveNext
Loop
rsM210.MoveFirst

'CUMULATIVA
'cM200 = "|" & "M200" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & Round(cVal210, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
'NAO CUMULATIVA
Set rsM200 = Db.OpenRecordset("select * from tbResumo_PIS where ano ='" & year(cDtIni) & "' and mes = " & month(cDtIni) & ";")

If rsM200!SALDO < 0 Then
cM200 = "|" & "M200" & "|" & Round(cVal210, 2) & "|" & Round(rsM200!CRED, 2) & "|" & Round(rsM200!CRED_MES_ANT, 2) & "|" & Round(rsM200!SALDO * -1, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
Else
    
    If rsM200!CRED > cVal210 Then
    cM200 = "|" & "M200" & "|" & Round(cVal210, 2) & "|" & Round(cVal210, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
    Else
        If rsM200!Saldo_Mes < 0 Then
        cM200 = "|" & "M200" & "|" & Round(cVal210, 2) & "|" & Round(rsM200!CRED, 2) & "|" & Round(Abs(rsM200!Saldo_Mes), 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
        Else
        cM200 = "|" & "M200" & "|" & Round(cVal210, 2) & "|" & Round(rsM200!CRED, 2) & "|" & Round(rsM200!CRED_MES_ANT, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
        End If
    End If
End If

cLinM = cLinM + 1
lM200 = lM200 + 1
Print #iArq, cM200
Else
End If

'Registro M205: Contribuição para o PIS/Pasep a Recolher–Detalhamento por Código de Receita
'cM205
'cM205 = "|" & "M205" & "|" & "12" & "|" & "067903" & "|" & Round(cVal210, 2) & "|"
If rsM200!SALDO < 0 Then
cM205 = "|" & "M205" & "|" & "08" & "|" & "067903" & "|" & Round(rsM200!SALDO * -1, 2) & "|"
cLinM = cLinM + 1
lM205 = lM205 + 1
Print #iArq, cM205
Else
End If


'cM210
'campos novos 01/01/19
'5-VL_AJUS_ACRESBC_PIS"
'6-VL_AJUS_REDUC_BC_PIS
'7-VL_BC_CONT_AJUS

Dim BasCalc210 As Double


Do Until rsM210.EOF

If rsM210!Aliq_PIS = 0 Then
BasCalc210 = 0
Else
BasCalc210 = Round(rsM210!ValorTot, 2)
End If

'cM210 = "|" & "M210" & "|" & "52" & "|" & Round(rsM210!ValorTot, 2) & "|" & Round(rsM210!ValorTot, 2) & "|" & "0" & "|" & "0" & "|" & Round(rsM210!ValorTot, 2) & "|" & rsM210!Aliq_PIS & "|" & "|" & "|" & Round(rsM210!Valor_PIS, 2) & "|" & "0" & "|" & "0" & "|" & "|" & "|" & Round(rsM210!Valor_PIS, 2) & "|"
If rsM210!Valor_PIS <> 0 Then
cM210 = "|" & "M210" & "|" & "02" & "|" & Round(rsM210!ValorTot, 2) & "|" & BasCalc210 & "|" & "0" & "|" & "0" & "|" & BasCalc210 & "|" & rsM210!Aliq_PIS & "|" & "|" & "|" & Round(rsM210!Valor_PIS, 2) & "|" & "0" & "|" & "0" & "|" & "|" & "|" & Round(rsM210!Valor_PIS, 2) & "|"
cLinM = cLinM + 1
lM210 = lM210 + 1
Print #iArq, cM210
Else: End If

rsM210.MoveNext
Loop


'M400: Receitas Isentas, não Alcançadas pela Incidência da Contribuição, Sujeitas a
'Alíquota Zero ou de Vendas com Suspensão – PIS/Pasep
Set rsM400 = Db.OpenRecordset("select CST_PIS, sum(ValorTot-VlrDesc) as Vlr_Item from tbVendas as q1 inner join tbVendasDet as q2 on q1.ID = q2.IDVenda " & _
"where DataEmissao >= #" & cDtIniVb & "# and DataEmissao <= #" & cDtFimVb & " 23:59:59" & "# " & _
"and CST_PIS in ('04','05','06','07','08','09') group by CST_PIS;")

Do Until rsM400.EOF
            cM400 = "|" & "M400" & "|" & rsM400!CST_PIS & "|" & rsM400!Vlr_Item & "|" & "4" & "|" & "Revenda de mercadoria adiq de terceiros sujeita a subs tributária e aliq monofasica" & "|"
            cLinM = cLinM + 1
            lm400 = lm400 + 1
            Print #iArq, cM400
            
            cM410 = "|" & "M410" & "|" & "427" & "|" & rsM400!Vlr_Item & "|" & "4" & "|" & "Revenda de mercadoria adiq de terceiros sujeita a subs tributária e aliq monofasica" & "|"
            cLinM = cLinM + 1
            lm410 = lm410 + 1
            Print #iArq, cM410
            
rsM400.MoveNext
Loop

If lm400 = "" Then
lm400 = 0
Else
End If
If lm410 = "" Then
lm410 = 0
Else
End If



'Registro M505: Detalhamento da Base de Calculo do Crédito Apurado no Período–Cofins
'APENAS PARA LUCRO REAL

Set rsM500 = Db.OpenRecordset("SELECT  Sum(tbComprasDet.ValorTot-tbComprasDet.VlrDesc) AS ValorTot, Sum(tbComprasDet.BaseCalc_PisCofins) AS BaseCalc_PisCofins, tbComprasDet.Aliq_COFINS, Sum(tbComprasDet.Valor_COFINS) AS Valor_COFINS, tbComprasDet.CST_COFINS " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE tbCompras.DataEmissao >= #" & cDtIniVb & "# And tbCompras.DataEmissao <= #" & cDtFimVb & " 23:59:59" & "# and LancFiscal = 'CREDITO' " & _
"GROUP BY tbComprasDet.Aliq_COFINS, tbComprasDet.CST_COFINS;")

If rsM500.EOF = True And rsM500.BOF = True Then
Else

Do Until rsM500.EOF


            If rsM500!Aliq_Cofins = 7.6 Or rsM500!Aliq_Cofins = 3 Then
            tAliq = 101
            Else
            tAliq = 102
            End If
 
            cM500 = "|" & "M500" & "|" & tAliq & "|" & "0" & "|" & rsM500!BaseCalc_PisCofins & "|" & rsM500!Aliq_Cofins & "|" & "|" & "|" & Round((rsM500!BaseCalc_PisCofins * rsM500!Aliq_Cofins) / 100, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & Round((rsM500!BaseCalc_PisCofins * rsM500!Aliq_Cofins) / 100, 2) & "|" & "0" & "|" & Round((rsM500!BaseCalc_PisCofins * rsM500!Aliq_Cofins) / 100, 2) & "|" & "0" & "|"
            cLinM = cLinM + 1
            lM500 = lM500 + 1
            Print #iArq, cM500



Set rsM505 = Db.OpenRecordset("SELECT IIf(IdFornecedor=1131,'04',IIf(MAT_PRIMA='SIM','02',IIf(REVENDA='SIM','01',IIf(EMBALAGEM='SIM','02')))) AS Natureza, Sum(tbComprasDet.ValorTot-tbComprasDet.VlrDesc) AS ValorTot, Sum(tbComprasDet.BaseCalc_PisCofins) AS BaseCalc_PisCofins, tbComprasDet.Aliq_COFINS, Sum(tbComprasDet.Valor_COFINS) AS Valor_COFINS, tbComprasDet.CST_COFINS " & _
"FROM (tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra) INNER JOIN tbCadProd ON tbComprasDet.IDProd = tbCadProd.IDProd " & _
"WHERE tbCompras.DataEmissao>=#" & cDtIniVb & "# And tbCompras.DataEmissao<=#" & cDtFimVb & " 23:59:59" & "# AND tbComprasDet.LancFiscal='CREDITO' AND tbComprasDet.Aliq_COFINS= " & Replace(rsM500!Aliq_Cofins, ",", ".") & " " & _
"GROUP BY IIf(IdFornecedor=1131,'04',IIf(MAT_PRIMA='SIM','02',IIf(REVENDA='SIM','01',IIf(EMBALAGEM='SIM','02')))), tbComprasDet.Aliq_COFINS, tbComprasDet.CST_COFINS;")


If rsM505.EOF = True And rsM505.BOF = True Then
Else

Do Until rsM505.EOF = True

            
            cM505 = "|" & "M505" & "|" & rsM505!NATUREZA & "|" & rsM505!CST_Cofins & "|" & Round(rsM505!BaseCalc_PisCofins, 2) & "|" & "0" & "|" & Round(rsM505!BaseCalc_PisCofins, 2) & "|" & Round(rsM505!BaseCalc_PisCofins, 2) & "|" & "|" & "|" & "Credito COFINS" & "|"
            cLinM = cLinM + 1
            lM505 = lM505 + 1
            Print #iArq, cM505


rsM505.MoveNext
Loop
End If

rsM500.MoveNext
Loop
End If





Dim cVal610 As Double

If rsM610.RecordCount > 0 Then

'Registro M600:  Consolidação da Contribuição para a Seguridade Social-Cofins do Período
Do Until rsM610.EOF = True
cVal610 = cVal610 + rsM610!Valor_Cofins
rsM610.MoveNext
Loop

rsM610.MoveFirst

Set rsM600 = Db.OpenRecordset("select * from tbResumo_COFINS where ano ='" & year(cDtIni) & "' and mes = " & month(cDtIni) & ";")

If rsM600!SALDO < 0 Then
cM600 = "|" & "M600" & "|" & Round(cVal610, 2) & "|" & Round(rsM600!CRED, 2) & "|" & Round(rsM600!CRED_MES_ANT, 2) & "|" & Round(rsM600!SALDO * -1, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
Else
    If rsM600!CRED > cVal610 Then
    cM600 = "|" & "M600" & "|" & Round(cVal610, 2) & "|" & Round(cVal610, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
    Else
    If rsM600!Saldo_Mes < 0 Then
    cM600 = "|" & "M600" & "|" & Round(cVal610, 2) & "|" & Round(rsM600!CRED, 2) & "|" & Round(Abs(rsM600!Saldo_Mes), 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
    Else
    cM600 = "|" & "M600" & "|" & Round(cVal610, 2) & "|" & Round(rsM600!CRED, 2) & "|" & Round(rsM600!CRED_MES_ANT, 2) & "|" & Round(rsM600!SALDO, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & 0 & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
    End If
    End If
End If

cLinM = cLinM + 1
lM600 = lM600 + 1
Print #iArq, cM600



'Registro M605: Cofins a Recolher–Detalhamento por Código de Receita
If rsM600!SALDO < 0 Then
cM605 = "|" & "M605" & "|" & "08" & "|" & "076003" & "|" & Round(rsM600!SALDO * -1, 2) & "|"
cLinM = cLinM + 1
lM605 = lM605 + 1
Print #iArq, cM605
Else
End If

Dim BasCalc610 As Double

Do Until rsM610.EOF

If rsM610!Aliq_Cofins = 0 Then
BasCalc610 = 0
Else
BasCalc610 = Round(rsM610!ValorTot, 2)
End If

If rsM610!Valor_Cofins <> 0 Then
cM610 = "|" & "M610" & "|" & "02" & "|" & Round(rsM610!ValorTot, 2) & "|" & BasCalc610 & "|" & "0" & "|" & "0" & "|" & BasCalc610 & "|" & rsM610!Aliq_Cofins & "|" & "|" & "|" & Round(rsM610!Valor_Cofins, 2) & "|" & "0" & "|" & "0" & "|" & "|" & "|" & Round(rsM610!Valor_Cofins, 2) & "|"
cLinM = cLinM + 1
lM610 = lM610 + 1
Print #iArq, cM610
Else: End If

rsM610.MoveNext
Loop


Set rsM800 = Db.OpenRecordset("select CST_Cofins, sum(ValorTot-VlrDesc) as Vlr_Item from tbVendas as q1 inner join tbVendasDet as q2 on q1.ID = q2.IDVenda " & _
"where DataEmissao >= #" & cDtIniVb & "# and DataEmissao <= #" & cDtFimVb & " 23:59:59" & "# " & _
"and CST_Cofins in ('04','05','06','07','08','09') group by CST_Cofins;")

Do Until rsM800.EOF
            cM800 = "|" & "M800" & "|" & rsM800!CST_Cofins & "|" & rsM800!Vlr_Item & "|" & "4" & "|" & "Revenda de mercadoria adiq de terceiros sujeita a subs tributária e aliq monofasica" & "|"
            cLinM = cLinM + 1
            lm800 = lm800 + 1
            Print #iArq, cM800
            
            cM810 = "|" & "M810" & "|" & "427" & "|" & rsM800!Vlr_Item & "|" & "4" & "|" & "Revenda de mercadoria adiq de terceiros sujeita a subs tributária e aliq monofasica" & "|"
            cLinM = cLinM + 1
            lm810 = lm810 + 1
            Print #iArq, cM810
            
rsM800.MoveNext
Loop
If lm800 = "" Then
lm800 = 0
Else
End If
If lm810 = "" Then
lm810 = 0
Else
End If




Else: End If

semdadosM:
'Registro M990: Encerramento do Bloco M
'cM990
cLinM = cLinM + 1
lM990 = lM990 + 1
cM990 = "|" & "M990" & "|" & cLinM & "|"
Print #iArq, cM990

'BLOCO P: Apuração da Contribuição Previdenciária Sobre a Receita Bruta (CPRB)
'Registro P001: Abertura do Bloco P
'cP001
cP001 = "|" & "P001" & "|" & "1" & "|"
cLinP = cLinP + 1
lP001 = lP001 + 1
Print #iArq, cP001

'Registro P990: Encerramento do Bloco P
'cP990
cLinP = cLinP + 1
lP990 = lP990 + 1
cP990 = "|" & "P990" & "|" & cLinP & "|"
Print #iArq, cP990


'BLOCO  1:  Complemento  da  Escrituração–Controle  de  Saldos  de  Créditos  e  deRetenções, Operações Extemporâneas e Outras Informações

'Creditos
Dim cAno As Integer
Dim cMes As Integer
Dim cTeste As Boolean


cAno = year(cDtIni)
cMes = month(cDtIni)
cMes = cMes - 1
If cMes < 1 Then
cMes = 12
cAno = cAno - 1
Else
End If


cTeste = False

Set RS1100 = Db.OpenRecordset("select * from tbResumo_PIS where ano ='" & year(cDtIni) & "' and mes = " & month(cDtIni) & ";")
If RS1100!CRED_MES_ANT > 0 Then
cTeste = True
Else
End If

Set RS1500 = Db.OpenRecordset("select * from tbResumo_COFINS where ano ='" & year(cDtIni) & "' and mes = " & month(cDtIni) & ";")
If RS1500!CRED_MES_ANT > 0 Then
cTeste = True
Else
End If

'Registro 1001: Abertura do Bloco 1
'c1001

If cTeste = True Then
    c1001 = "|" & "1001" & "|" & "0" & "|"
    cLin1 = cLin1 + 1
    l1001 = l1001 + 1
    Print #iArq, c1001

Else
    c1001 = "|" & "1001" & "|" & "1" & "|"
    cLin1 = cLin1 + 1
    l1001 = l1001 + 1
    Print #iArq, c1001
End If


'Registro 1100: Controle de Créditos Fiscais–PIS/Pasep - Anterior
'Tem que fazer para lançar o crédito do período anterior
'http://www.e-auditoria.com.br/publicacoes/artigos/como-aproveitar-saldos-de-creditos-fiscais-de-pispasep-e-cofins-de-periodos-anteriores-ao-da-escrituracao/
If RS1100!CRED_MES_ANT > 0 Then
    
    c1100 = "|" & "1100" & "|" & Format(cMes, "00") & cAno & "|" & "01" & "|" & "|" & "101" & "|" & Replace(Round(RS1100!CRED_MES_ANT, 2), ".", ",") & "|" & "|" & Replace(Round(RS1100!CRED_MES_ANT, 2), ".", ",") & "|" & "0" & "|" & "|" & "|" & Replace(Round(RS1100!CRED_MES_ANT, 2), ".", ",") & "|" & Replace(Round(Abs(RS1100!Saldo_Mes), 2), ".", ",") & "|" & "|" & "|" & "|" & "|" & Replace(Round(Abs(RS1100!SALDO), 2), ".", ",") & "|"
    
    cLin1 = cLin1 + 1
    l1100 = l1100 + 1
    Print #iArq, c1100
Else
l1100 = 0
End If

'Registro 1500: Controle de Créditos Fiscais–Cofins - Anterior
'Tem que fazer para lançar o crédito do período anterior
If RS1500!CRED_MES_ANT > 0 Then
    c1500 = "|" & "1500" & "|" & Format(cMes, "00") & cAno & "|" & "01" & "|" & "|" & "101" & "|" & Replace(Round(RS1500!CRED_MES_ANT, 2), ".", ",") & "|" & "|" & Replace(Round(RS1500!CRED_MES_ANT, 2), ".", ",") & "|" & "0" & "|" & "|" & "|" & Replace(Round(RS1500!CRED_MES_ANT, 2), ".", ",") & "|" & Replace(Round(Abs(RS1500!Saldo_Mes), 2), ".", ",") & "|" & "|" & "|" & "|" & "|" & Replace(Round(Abs(RS1500!SALDO), 2), ".", ",") & "|"
    cLin1 = cLin1 + 1
    l1500 = l1500 + 1
    Print #iArq, c1500
Else
l1500 = 0
End If


'Registro 1900: Consolidação dos Documentos Emitidos no Período por Pessoa Jurídica Submetida ao Regime de Tributação Com Base no Lucro Presumido–Regime de Caixa ou de Competência
'Parece que é apenas para não cumulativo

'Registro 1990: Encerramento do Bloco 1
'c1990
l1001 = l1001 + 1
cLin1 = cLin1 + 1
c1990 = "|" & "1990" & "|" & cLin1 & "|"
Print #iArq, c1990


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
'c0100
c9900 = "|" & "9900" & "|" & "0100" & "|" & l0100 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0110
c9900 = "|" & "9900" & "|" & "0110" & "|" & l0110 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0140
c9900 = "|" & "9900" & "|" & "0140" & "|" & l0140 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0150
c9900 = "|" & "9900" & "|" & "0150" & "|" & l0150 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0190
c9900 = "|" & "9900" & "|" & "0190" & "|" & l0190 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0200
c9900 = "|" & "9900" & "|" & "0200" & "|" & l0200 & "|"
Print #iArq, c9900
l99 = l99 + 1
If l0400 = "" Then
Else
'c0400
c9900 = "|" & "9900" & "|" & "0400" & "|" & l0400 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'c0500
c9900 = "|" & "9900" & "|" & "0500" & "|" & l0500 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0600
c9900 = "|" & "9900" & "|" & "0600" & "|" & l0600 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c0990
c9900 = "|" & "9900" & "|" & "0990" & "|" & l0990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO A:
'cA001
c9900 = "|" & "9900" & "|" & "A001" & "|" & lA001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cA990
c9900 = "|" & "9900" & "|" & "A990" & "|" & lA990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO C:
'cC001
c9900 = "|" & "9900" & "|" & "C001" & "|" & lC001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cC010
c9900 = "|" & "9900" & "|" & "C010" & "|" & lC010 & "|"
Print #iArq, c9900
l99 = l99 + 1
If lC100 = "" Then
Else
'cC100
c9900 = "|" & "9900" & "|" & "C100" & "|" & lC100 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cC170
c9900 = "|" & "9900" & "|" & "C170" & "|" & lC170 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'cC180
If lC180 = "" Then
Else
c9900 = "|" & "9900" & "|" & "C180" & "|" & lC180 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'cC181
If lC181 = "" Then
Else
c9900 = "|" & "9900" & "|" & "C181" & "|" & lC181 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'cC185
If lC185 = "" Then
Else
c9900 = "|" & "9900" & "|" & "C185" & "|" & lC185 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
If rsEnergia.RecordCount > 0 Then
'cC500
c9900 = "|" & "9900" & "|" & "C500" & "|" & lC500 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cC501
c9900 = "|" & "9900" & "|" & "C501" & "|" & lC501 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cC505
c9900 = "|" & "9900" & "|" & "C505" & "|" & lC505 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cC860
c9900 = "|" & "9900" & "|" & "C860" & "|" & lC860 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cC870
c9900 = "|" & "9900" & "|" & "C870" & "|" & lC870 & "|"
Print #iArq, c9900
l99 = l99 + 1
Else: End If
'cC990
c9900 = "|" & "9900" & "|" & "C990" & "|" & lC990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO D:
'cD001
c9900 = "|" & "9900" & "|" & "D001" & "|" & lD001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cD010
If lD010 = "" Then
Else
c9900 = "|" & "9900" & "|" & "D010" & "|" & lD010 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'cD100
If lD100 = "" Then
Else
c9900 = "|" & "9900" & "|" & "D100" & "|" & lD100 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If


'cD101
If lD101 = "" Then
Else
c9900 = "|" & "9900" & "|" & "D101" & "|" & lD101 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'cD105
If lD105 = "" Then
Else
c9900 = "|" & "9900" & "|" & "D105" & "|" & lD105 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If

'cD990
c9900 = "|" & "9900" & "|" & "D990" & "|" & lD990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO F:
'cF001
c9900 = "|" & "9900" & "|" & "F001" & "|" & lF001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cF010
c9900 = "|" & "9900" & "|" & "F010" & "|" & lF010 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cF550
If lF550 = "" Then
Else
c9900 = "|" & "9900" & "|" & "F550" & "|" & lF550 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'cF990
c9900 = "|" & "9900" & "|" & "F990" & "|" & lF990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO M:
'cM001
c9900 = "|" & "9900" & "|" & "M001" & "|" & lM001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cM100
c9900 = "|" & "9900" & "|" & "M100" & "|" & lM100 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cM105
c9900 = "|" & "9900" & "|" & "M105" & "|" & lM105 & "|"
Print #iArq, c9900
l99 = l99 + 1

'If rsVenda.RecordCount > 0 Then
'cM200
c9900 = "|" & "9900" & "|" & "M200" & "|" & lM200 & "|"
Print #iArq, c9900
l99 = l99 + 1


If rsM210.RecordCount > 0 Then
    If rsM200!SALDO < 0 Then
'cM205
c9900 = "|" & "9900" & "|" & "M205" & "|" & lM205 & "|"
Print #iArq, c9900
l99 = l99 + 1
    Else
    End If

'cM210
c9900 = "|" & "9900" & "|" & "M210" & "|" & lM210 & "|"
Print #iArq, c9900
l99 = l99 + 1
Else: End If

'cM400
c9900 = "|" & "9900" & "|" & "M400" & "|" & lm400 & "|"
Print #iArq, c9900
l99 = l99 + 1

'cM410
c9900 = "|" & "9900" & "|" & "M410" & "|" & lm410 & "|"
Print #iArq, c9900
l99 = l99 + 1


'cM500
c9900 = "|" & "9900" & "|" & "M500" & "|" & lM500 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cM505
c9900 = "|" & "9900" & "|" & "M505" & "|" & lM505 & "|"
Print #iArq, c9900
l99 = l99 + 1

'cM600
c9900 = "|" & "9900" & "|" & "M600" & "|" & lM600 & "|"
Print #iArq, c9900
l99 = l99 + 1

If rsM610.RecordCount > 0 Then
    If rsM600!SALDO < 0 Then
'cM605
c9900 = "|" & "9900" & "|" & "M605" & "|" & lM605 & "|"
Print #iArq, c9900
l99 = l99 + 1
    Else
    End If
'cM610
c9900 = "|" & "9900" & "|" & "M610" & "|" & lM610 & "|"
Print #iArq, c9900
l99 = l99 + 1
Else: End If

'cM800
c9900 = "|" & "9900" & "|" & "M800" & "|" & lm800 & "|"
Print #iArq, c9900
l99 = l99 + 1

'cM810
c9900 = "|" & "9900" & "|" & "M810" & "|" & lm810 & "|"
Print #iArq, c9900
l99 = l99 + 1


'Else
'End If
'cM990
c9900 = "|" & "9900" & "|" & "M990" & "|" & lM990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO P:
'cP001
c9900 = "|" & "9900" & "|" & "P001" & "|" & lP001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'cP990
c9900 = "|" & "9900" & "|" & "P990" & "|" & lP990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO  1:
'c1001
c9900 = "|" & "9900" & "|" & "1001" & "|" & l1001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'c1100
If l1100 = 0 Then
Else
c9900 = "|" & "9900" & "|" & "1100" & "|" & l1100 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'c1500
If l1500 = 0 Then
Else
c9900 = "|" & "9900" & "|" & "1500" & "|" & l1500 & "|"
Print #iArq, c9900
l99 = l99 + 1
End If
'c1990
c9900 = "|" & "9900" & "|" & "1990" & "|" & "1" & "|"
Print #iArq, c9900
l99 = l99 + 1

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
'totalizador 9900
c9900 = "|" & "9900" & "|" & "9900" & "|" & l99 & "|"
Print #iArq, c9900


'totalizador 9900
'c9900 = "|" & "9900" & "|" & "9900" & "|" & "1" & "|"
'Print #iArq, c9900

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
l99 = l99 + 1
c9990 = "|" & "9990" & "|" & l99 & "|"
Print #iArq, c9990





cTotal9999 = l0000 + l0001 + l0100 + l0110 + l0140 + l0150 + l0190 + l0200 + l0400 + l0500 + l0600 + l0990 + lA001 + lA990 + lC001 + lC010 + lC100 + lC170 + lC180 + lC181 + lC185 + lC500 + lC501 + lC505 + lC860 + lC870 + lC990 + lC990 + lD001 + lD010 + lD100 + lD101 + lD105 + lD990 + lF001 + lF010 + lF550 + lF990 + lM001 + lM100 + lM105 + lM200 + lM205 + lM210 + lm400 + lm410 + lM500 + lM505 + lM600 + lM605 + lM610 + lm800 + lm810 + lM990 + lP001 + lP990 + l1001 + l1100 + l1500 + l1990 + l9001 + l9900
c9999 = "|" & "9999" & "|" & cTotal9999 + l99 - 3 & "|"
Print #iArq, c9999


Close #iArq
'DoCmd.setwarnings (True)
MsgBox ("Arquivo Gerado")

Call DisconnectFromDataBase

End Function



