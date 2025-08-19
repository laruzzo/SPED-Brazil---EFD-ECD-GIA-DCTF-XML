Attribute VB_Name = "modGeraArquivo_EFD_ICMS_IPI"
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


Public Function Gerar_EFD_ICMS_IPI(cDtIni As String, cDtFim As String, clocal As String, cDtINI_Contabil As String, cIDIventario As String)
'VERSÃO 15 - 15/02/2021
Call ConnectToDataBase
'ARQUIVO EFD
'DoCmd.setwarnings (False)

'EXPORTAR ARQUIVO TXT
Dim iArq As Long
iArq = FreeFile

Open clocal & "\EFD_ICMS_IPI_" & month(cDtIni) & "_" & year(cDtIni) & ".txt" For Output As iArq

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
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIniVb & "# And (tbVendas.DataEmissao) <= #" & cDtFimVb & " 23:59:59" & "#)) " & _
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
'Duplicou, fiz ajuste para eliminar duplicado antes de adicionar
'strSQL = ("insert into tbFornecedor_Ativo_temp " & _
'"SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email " & _
'"FROM tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
'"WHERE (((tbCompras.DataEmissao) >= '" & cDtIni & "' And (tbCompras.DataEmissao) <= '" & cDtFim & "')) " & _
'"GROUP BY tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email;")
strSQL = ("insert into tbFornecedor_Ativo_temp " & _
"SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email " & _
"FROM tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor LEFT OUTER JOIN tbFornecedor_Ativo_temp ON  tbFornecedor.IDFor = tbFornecedor_Ativo_temp.IDFor " & _
"WHERE tbCompras.DataEmissao >= '" & cDtIni & "' And tbCompras.DataEmissao <= '" & cDtFim & "' and tbFornecedor_Ativo_temp.IDFor is null " & _
"GROUP BY tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email")
Conn.Execute strSQL


'insere transportadoras
strSQL = ("INSERT INTO tbFornecedor_Ativo_temp ( IDFor, Tipo, Cnpj, RazaoSocial, IE, CRT, CEP, Logradouro, Nro, Compl, Bairro, UF, cod_Municipio, Municipio, Pais, Fone, Email ) " & _
"SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email " & _
"FROM tbFornecedor INNER JOIN tbTransportes ON tbFornecedor.IDFor = tbTransportes.ID_Emit " & _
"where tbTransportes.LancFiscal = 'CREDITO' " & _
"GROUP BY tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email, tbTransportes.DataEmissao " & _
"HAVING tbTransportes.DataEmissao>='" & cDtIni & "' And tbTransportes.DataEmissao<='" & cDtFim & "';")
Conn.Execute strSQL

'REMOVE A IMOBILIARIA QUE NÃO DÁ CREDITO DE ICMS
strSQL = ("DELETE FROM tbFornecedor_Ativo_temp WHERE tbFornecedor_Ativo_temp.IDFor=1136")
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
"WHERE (((tbCompras.DataEmissao) >= '" & cDtIni & "' And (tbCompras.DataEmissao) <= ' " & cDtFim & " ')) " & _
"GROUP BY tbCadProd.IDProd;")
Conn.Execute strSQL

Dim rsCFOP As DAO.Recordset
'Set rsCFOP = db.OpenRecordset("SELECT tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
'"FROM tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbComprasDet INNER JOIN tbCadProd_Ativo_temp ON tbComprasDet.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
'"GROUP BY tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbCompras.IdFornecedor " & _
'"HAVING (((tbCompras.IdFornecedor)<>1131)); " & _
'"UNION " & _
'"SELECT tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC FROM tbVendasDet INNER JOIN tbCadProd_Ativo_temp ON tbVendasDet.IDProd = tbCadProd_Ativo_temp.IDProd " & _
'"GROUP BY tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
'"HAVING (((tbVendasDet.CFOP_ESCRITURADA) Is Not Null));")

Set rsCFOP = Db.OpenRecordset("SELECT tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
"FROM tbCompras INNER JOIN (tbComprasDet INNER JOIN tbCadProd_Ativo_temp ON tbComprasDet.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.IdFornecedor) <> 1131)) " & _
"GROUP BY tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC " & _
"UNION " & _
"SELECT tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC " & _
"FROM tbVendas INNER JOIN (tbVendasDet INNER JOIN tbCadProd_Ativo_temp ON tbVendasDet.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"GROUP BY tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, tbVendas.TipoNF " & _
"HAVING (((tbVendasDet.CFOP_ESCRITURADA) Is Not Null) AND ((tbVendas.TipoNF)='0-ENTRADA'));")

'tira vendas

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

'CONSUMO DE INDUSTRIALIZAÇÃO
strSQL = ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
"SELECT tb_Registro_Consumo.ID_Produto " & _
"FROM (tb_Registro_Envase INNER JOIN tb_Registro_Consumo ON tb_Registro_Envase.ID = tb_Registro_Consumo.ID_Lote) LEFT JOIN tbCadProd_Ativo_temp ON tb_Registro_Consumo.ID_Produto = tbCadProd_Ativo_temp.IDProd " & _
"WHERE (((tb_Registro_Envase.DATA)>='" & cDtIni & "' And (tb_Registro_Envase.DATA)<='" & cDtFim & "') AND ((tbCadProd_Ativo_temp.IDProd) Is Null));")
Conn.Execute strSQL

'REGISTRO ENVASE
strSQL = ("INSERT INTO tbCadProd_Ativo_temp ( IDProd ) " & _
"SELECT tb_Registro_Envase.ID_Produto " & _
"FROM tb_Registro_Envase LEFT JOIN tbCadProd_Ativo_temp ON tb_Registro_Envase.ID_Produto = tbCadProd_Ativo_temp.IDProd " & _
"WHERE (((tb_Registro_Envase.DATA)>='" & cDtIni & "' And (tb_Registro_Envase.DATA)<='" & cDtFim & "') AND ((tbCadProd_Ativo_temp.IDProd) Is Null));")
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




If CDate(cDtIni) <= #12/31/2017# Then
cCodVer = "011"
Else
cCodVer = "012"
End If

If CDate(cDtIni) >= #1/1/2019# Then
cCodVer = "015"
Else
End If

If CDate(cDtIni) >= #1/1/2022# Then
cCodVer = "016"
Else
End If

If CDate(cDtIni) >= #1/1/2023# Then
cCodVer = "017"
Else
End If


cCodFin = "0"
cPerfil = "A"
cAtividade = "0"

clintot = 0


Dim c0000 As String
Dim c0001 As String
Dim c0002 As String
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

'REGISTRO 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA ENTIDADE
'REGISTRO 0001: ABERTURA DO BLOCO 0
'REGISTRO 0002: CLASSIFICAÇÃO   DO   ESTABELECIMENTO   INDUSTRIAL   OUEQUIPARADO A INDUSTRIAL
'REGISTRO 0005: DADOS COMPLEMENTARES DA ENTIDADE
'REGISTRO 0015: DADOS DO CONTRIBUINTE SUBSTITUTO OU RESPONSÁVEL PELO ICMS DESTINO
'REGISTRO 0100: DADOS DO CONTABILISTA
'REGISTRO 0150: TABELA DE CADASTRO DO PARTICIPANTE
'REGISTRO 0175: ALTERAÇÃO DA TABELA DE CADASTRO DE PARTICIPANTE
'REGISTRO 0190: IDENTIFICAÇÃO DAS UNIDADES DE MEDIDA
'REGISTRO 0200: TABELA DE IDENTIFICAÇÃO DO ITEM (PRODUTO E SERVIÇOS)
'REGISTRO 0205: ALTERAÇÃO DO ITEM
'REGISTRO 0210: CONSUMO ESPECÍFICO PADRONIZADO
'REGISTRO 0220: FATORES DE CONVERSÃO DE UNIDADES
'REGISTRO 0300: CADASTRO DE BENS OU COMPONENTES DO ATIVO IMOBILIZADO
'REGISTRO 0305: INFORMAÇÃO SOBRE A UTILIZAÇÃO DO BEM
'REGISTRO 0400: TABELA DE NATUREZA DA OPERAÇÃO/PRESTAÇÃO
'REGISTRO 0450: TABELA DE INFORMAÇÃO COMPLEMENTAR DO DOCUMENTO FISCAL
'REGISTRO 0500: PLANO DE CONTAS CONTÁBEIS
'REGISTRO 0600: CENTRO DE CUSTOS
'REGISTRO 0990: ENCERRAMENTO DO BLOCO 0



'0000
c0000 = "|" & "0000" & "|" & cCodVer & "|" & cCodFin & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|" & rsEmpresa!RazaoSocial & "|" & rsEmpresa!CNPJ & "|" & "|" & rsEmpresa!UF & "|" & rsEmpresa!IE & "|" & rsEmpresa!Cidade_IBGE & "|" & rsEmpresa!IM & "|" & "|" & cPerfil & "|" & cAtividade & "|"
l0000 = 1
clintot = clintot + 1
Print #iArq, c0000

c0001 = "|" & "0001" & "|" & "0" & "|"
clintot = clintot + 1
l0001 = 1
Print #iArq, c0001

c0002 = "|" & "0002" & "|" & "00" & "|"
clintot = clintot + 1
l0002 = 1
Print #iArq, c0002


c0005 = "|" & "0005" & "|" & rsEmpresa!NomeFantasia & "|" & rsEmpresa!CEP & "|" & rsEmpresa!Logradouro & "|" & rsEmpresa!Num & "|" & rsEmpresa!compl & "|" & rsEmpresa!Bairro & "|" & rsEmpresa!Fone & "||" & rsEmpresa!Email & "|"
clintot = clintot + 1
l0005 = 1
Print #iArq, c0005

'c0015 Omitido
c0100 = "|" & "0100" & "|" & rsContador!NomeContador & "|" & rsContador!CPFContador & "|" & rsContador!CRCContador & "|" & rsContador!CNPJEscritorio & "|" & rsContador!CEPEscritorio & "|" & rsContador!ENDEscritorio & "|" & rsContador!NumeroEscritorio & "|" & rsContador!ComplEscritorio & "|" & rsContador!BairroEscritorio & "|" & rsContador!TelefoneEscritorio & "||" & rsContador!EmailEscritorio & "|" & rsContador!CodMunicipioEscritorio & "|"
clintot = clintot + 1
l0100 = 1
Print #iArq, c0100


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

'MsgBox (c0150)
 'c0175 Omitido   - Registro de Alteração de cadastro
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
 c0200 = "|" & "0200" & "|" & rsCadProd!IDProd & "|" & rsCadProd!DescProd & "|" & rsCadProd!EAN & "|" & "|" & rsCadProd!Unid & "|" & cTipoItem & "|" & rsCadProd!NCM & "||" & "|||" & "|"
 Print #iArq, c0200
 rsCadProd.MoveNext
 l0200 = l0200 + 1
 clintot = clintot + 1
 End If
 Loop
 
 'c0205 Omitido
 'c0206 Omitido
 'c0210 Omitido
 'c0220 Omitido
 
 'c0300
  'imobilizado Bem principal de componentes
strSQL = ("INSERT INTO tbCadProd_Ativo_temp (IDPROD) " & _
"select CodBem from " & _
"(SELECT tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.Status, tbCadProd_Ativo_temp.IDProd FROM tbCadProd_Ativo_temp right outer JOIN tbImobilizadoCadastro ON tbCadProd_Ativo_temp.IDProd = tbImobilizadoCadastro.CodBem " & _
"GROUP BY tbImobilizadoCadastro.CodBem " & _
"HAVING tbCadProd_Ativo_temp.IDProd Is  Null and tbImobilizadoCadastro.CodBem is not null ) as q1;")
Conn.Execute strSQL

'deleta exauridos
strSQL = ("delete tbCadProd_Ativo_temp from tbCadProd_Ativo_temp inner join (select * from tbimobilizadocadastro where status = 'EXAURIDO') AS Q2 ON tbCadProd_Ativo_temp.IDProd = q2.IDProd;")
Conn.Execute strSQL

'DoCmd.RunSQL ("INSERT INTO tbCadProd_Ativo_temp ( IDPROD ) " & _
'"SELECT tbImobilizadoCadastro.CodBem " & _
'"FROM (tbCadProd_Ativo_temp INNER JOIN tbImobilizadoCadastro ON tbCadProd_Ativo_temp.IDProd = tbImobilizadoCadastro.IDProd) INNER JOIN tbImobilizado ON tbImobilizadoCadastro.IDProd = tbImobilizado.IDProd " & _
'"WHERE (((tbImobilizado.Valor_ICMS) > 0) And ((tbImobilizado.ANO) = " & Year(cDtINI) & ") And ((tbImobilizado.MES) = " & Month(cDtINI) & ")) " & _
'"GROUP BY tbImobilizadoCadastro.CodBem " & _
'"HAVING (((tbImobilizadoCadastro.CodBem) Is Not Null));")


Dim rsImob As DAO.Recordset
'com excessos
'Set rsImob = DB.OpenRecordset("SELECT tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
'"FROM tbPlanoContasContabeis INNER JOIN (tbImobilizadoCadastro INNER JOIN tbCadProd_Ativo_temp ON tbImobilizadoCadastro.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta " & _
'"GROUP BY tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza;")
'sem excessos
'Set rsImob = DB.OpenRecordset("SELECT tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
'"FROM tbPlanoContasContabeis INNER JOIN ((tbImobilizado INNER JOIN tbCadProd_Ativo_temp ON tbImobilizado.IDProd = tbCadProd_Ativo_temp.IDProd) INNER JOIN tbImobilizadoCadastro ON tbImobilizado.IDProd = tbImobilizadoCadastro.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta " & _
'"GROUP BY tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza;")

'SEM EXCESSOS E COM O BEM FALTANTE
Set rsImob = Db.OpenRecordset("SELECT tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
"FROM tbPlanoContasContabeis INNER JOIN ((tbImobilizado INNER JOIN tbCadProd_Ativo_temp ON tbImobilizado.IDProd = tbCadProd_Ativo_temp.IDProd) INNER JOIN tbImobilizadoCadastro ON tbImobilizado.IDProd = tbImobilizadoCadastro.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta " & _
"GROUP BY tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
"UNION " & _
"select distinct q2.* from " & _
"(SELECT tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
"FROM tbPlanoContasContabeis INNER JOIN ((tbImobilizado INNER JOIN tbCadProd_Ativo_temp ON tbImobilizado.IDProd = tbCadProd_Ativo_temp.IDProd) INNER JOIN tbImobilizadoCadastro ON tbImobilizado.IDProd = tbImobilizadoCadastro.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta " & _
"GROUP BY tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
"HAVING (((tbImobilizadoCadastro.Bem_Componente) = 'COMP')) " & _
") as q1 " & _
"INNER Join " & _
"(SELECT tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
"FROM tbPlanoContasContabeis INNER JOIN (tbImobilizadoCadastro INNER JOIN tbCadProd_Ativo_temp ON tbImobilizadoCadastro.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta " & _
"GROUP BY tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza " & _
"HAVING(tbImobilizadoCadastro.Bem_Componente) = 'BEM' " & _
") AS Q2 " & _
"ON Q1.CodBem = Q2.IdProd")


 cBemComp = ""
 c0300 = ""
 Dim cId As Integer
 cId = 0
Do Until rsImob.EOF
 
 Select Case rsImob!Bem_Componente
 Case Is = "BEM"
 cBemComp = "1"
  Case Is = "COMP"
 cBemComp = "2"
 Case Else
 cBemComp = "1"
 End Select
 
 Select Case rsImob!Centro_Custo
 Case Is = "PROD"
 cBemCC = "3"
 cBemCCDEsc = "Area Produtiva"
  Case Is = "ADM"
 cBemCC = "5"
 cBemCCDEsc = "Area Administrativa"
 Case Else
 cBemCC = "3"
 cBemCCDEsc = "Area Produtiva"
 End Select
 cId = rsImob!IDProd
 c0300 = "|" & "0300" & "|" & rsImob!IDProd & "|" & cBemComp & "|" & rsImob!DESCRICAO & "|" & rsImob!CodBem & "|" & rsImob!ID_Conta & "|" & rsImob!Nr_Parcelas & "|"
 Print #iArq, c0300
 clintot = clintot + 1
 l0300 = l0300 + 1
 c0305 = "|" & "0305" & "|" & cBemCC & "|" & cBemCCDEsc & "|" & rsImob!Nr_Parcelas & "|"
 Print #iArq, c0305
 clintot = clintot + 1
 l0305 = l0305 + 1
 rsImob.MoveNext
 If rsImob.EOF = True Then
 Else
 If cId = rsImob!IDProd Then
 rsImob.MoveNext
 Else
 End If
 End If
 Loop
 
'c0305
c0305 = ""
'rsImob.MoveFirst
'Do Until rsImob.EOF
'Select Case rsImob!Centro_Custo
'Case Is = "PROD"
'cCusto = "3"
'cFunc = "Uso para atividade fabricação de cerveja"
'Case Is = "ADM"
'cCusto = "5"
'cFunc = "Uso para atividades administrativas da empresa"
'Case Else
'cCusto = "3"
'End Select

'Administrativa
'c0305 = "|" & "0305" & "|" & "2" & "|" & "Area Administrativa" & "|" & "48" & "|"
'Print #iArq, c0305
'clintot = clintot + 1
'Produtiva
'c0305 = "|" & "0305" & "|" & "3" & "|" & "Area Produtiva" & "|" & "48" & "|"
'Print #iArq, c0305
'clintot = clintot + 1


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

'c0450 Omitido
'c0460 Omitido

'c0500
c0500 = ""
Do Until rsContas.EOF = True
c0500 = "|" & "0500" & "|" & cSTR_DtINI & "|" & rsContas!Cod_Natureza & "|" & rsContas!Cod_Indicador & "|" & "1" & "|" & rsContas!ID & "|" & rsContas!Desc_CodNatureza & "|"
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

'BLOCO B - A partir 01/2019 BLOCO B: ESCRITURAÇÃO E APURAÇÃO DO ISS
'B001
cB001 = "|" & "B001" & "|" & "1" & "|"
cLinB = cLinB + 1
lB001 = lB001 + 1
Print #iArq, cB001
'B001

'B990
cLinB = cLinB + 1
lB990 = lB990 + 1
cB990 = "|" & "B990" & "|" & cLinB & "|"
Print #iArq, cB990
'B990

'BLOCO B

    
'C - Documentos Fiscais I  – Mercadorias (ICMS/IPI)
Dim cC001 As String
Dim cC100 As String


Dim cLinC As Integer


Dim rsCompra As DAO.Recordset
'Set rsCompra = DB.OpenRecordset("select * from tbCompras where dataemissao >= #" & cDtINI & "# and dataemissao <= #" & cDtFIM & "# and IdFornecedor <> 1131")
Set rsCompra = Db.OpenRecordset("SELECT tbCompras.ID, tbCompras.IdFornecedor, tbCompras.Serie, tbCompras.NumNF, tbCompras.ChaveNF, tbCompras.DataEmissao, tbCompras.VlrTOTALNF, tbCompras.VlrDesconto, tbCompras.VlrTotalProdutos, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.BaseCalculo,0)) AS BaseCalculo, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_ICMS,0)) AS Valor_ICMS, Sum(tbComprasDet.BaseCalc_ST) AS BaseCalc_ST, Sum(tbComprasDet.Valor_ICMS_ST) AS Valor_ICMS_ST, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_IPI,0)) AS Valor_IPI, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_PIS,0)) AS Valor_PIS, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_Cofins,0)) AS Valor_Cofins " & _
"FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"GROUP BY tbCompras.ID, tbCompras.IdFornecedor, tbCompras.Serie, tbCompras.NumNF, tbCompras.ChaveNF, tbCompras.DataEmissao, tbCompras.VlrTOTALNF, tbCompras.VlrDesconto, tbCompras.VlrTotalProdutos, tbCompras.dataemissao " & _
"HAVING tbCompras.dataemissao>= #" & cDtIni & "# and dataemissao <= #" & cDtFim & "# and IdFornecedor <> 1131;")


Dim rsVenda As DAO.Recordset
Set rsVenda = Db.OpenRecordset("select * from tbVendas where dataemissao >= #" & cDtIni & "# and dataemissao <= #" & cDtFim & "# and tipoNF = '1-SAIDA' and [Status] = 'ATIVO' and NatOperacao <> 'Venda Cupom Fiscal SAT'")

Dim rsCIAP As DAO.Recordset
Set rsCIAP = Db.OpenRecordset("select * from tbVendas where dataemissao >= #" & cDtIni & "# and dataemissao <= #" & cDtFim & "# and tipoNF = '0-ENTRADA' and [Status] = 'ATIVO'")
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
Dim rsCIAPDet As DAO.Recordset


Dim rsRegSaida As DAO.Recordset
Set rsRegSaida = Db.OpenRecordset("SELECT tbvendas.NatOperacao, Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbvendasdet.CFOP_ESCRITURADA AS CFOP, tbvendasdet.CFOP_ESC_DESC AS [CFOP Desc], tbvendasdet.lancfiscal AS [Lanc Fiscal], Sum(tbvendasdet.ValorTot) AS [Valor Contabil], Sum(tbvendasdet.BaseCalculo) AS [Base de Calculo], Sum(tbvendasdet.Valor_ICMS) AS ICMS, Sum(tbvendasdet.Valor_IPI) AS IPI, Sum(tbvendasdet.Valor_PIS) AS PIS, Sum(tbvendasdet.Valor_Cofins) AS Cofins, Sum(tbvendasdet.Valor_ICMS_ST) AS [ICMS ST], Sum(tbvendasdet.BaseCalc_ST) AS BaseCalc_ST, tbvendasdet.CST, tbvendasdet.CST_DESC, tbvendasdet.Aliq_ICMS " & _
"FROM (tbCliente INNER JOIN tbvendas ON tbCliente.IDCliente = tbvendas.Idcliente) INNER JOIN (tbCadProd INNER JOIN tbvendasdet ON (tbCadProd.IDProd = tbvendasdet.IDProd) AND (tbCadProd.IDProd = tbvendasdet.IDProd)) ON tbvendas.ID = tbvendasdet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIni & "# And (tbVendas.DataEmissao) <= #" & cDtFim & "#)) " & _
"GROUP BY tbvendas.NatOperacao, Year(DataEmissao), Month(DataEmissao), tbvendasdet.CFOP_ESCRITURADA, tbvendasdet.CFOP_ESC_DESC, tbvendasdet.lancfiscal, tbvendas.TipoNF, tbvendasdet.CST, tbvendasdet.CST_DESC, tbvendasdet.Aliq_ICMS " & _
"HAVING tbvendas.TipoNF = '1-SAIDA' and tbVendas.NatOperacao <> 'Venda Cupom Fiscal SAT' " & _
"ORDER BY Year(DataEmissao), Month(DataEmissao);")

Dim rsRegEntr As DAO.Recordset

Dim rsEnergia As DAO.Recordset
Set rsEnergia = Db.OpenRecordset("SELECT * FROM tbCompras WHERE (((tbCompras.IdFornecedor)=1131) AND ((tbCompras.DataEmissao)>= #" & cDtIni & "#  And (tbCompras.DataEmissao)<= #" & cDtFim & "#));")

Dim rsEnergiaDet As DAO.Recordset
Set rsEnergiaDet = Db.OpenRecordset("SELECT tbComprasDet.* FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra WHERE (((tbCompras.IdFornecedor)=1131) AND ((tbCompras.DataEmissao)>=#" & cDtIni & "# And (tbCompras.DataEmissao)<=#" & cDtFim & "#));")

Dim rsEnergiaInjetada As DAO.Recordset
Set rsEnergiaInjetada = Db.OpenRecordset("SELECT * FROM tbEnergia where ANO =  year('" & cDtIni & "') and MES =  month('" & cDtIni & "')")


cLinC = 0

'REGISTRO C001: ABERTURA DO BLOCO C
'REGISTRO C100: NOTA FISCAL
'REGISTRO C101: INFORMAÇÃO COMPLEMENTAR DOS DOCUMENTOS FISCAIS QUANDO DAS OPERAÇÕES INTERESTADUAIS DESTINADAS A CONSUMIDOR FINAL NÃO CONTRIBUINTE EC 87/15 (CÓDIGO 55)
'REGISTRO C105: OPERAÇÕES COM ICMS ST RECOLHIDO PARA UF DIVERSA DO DESTINATÁRIO DO DOCUMENTO FISCAL (CÓDIGO 55)
'REGISTRO C110: INFORMAÇÃO COMPLEMENTAR DA NOTA FISCAL (CÓDIGO 01,1B, 04 e 55).
'REGISTRO C170: ITENS DO DOCUMENTO (CÓDIGO 01, 1B, 04 e 55).
'REGISTRO C190: REGISTRO ANALÍTICO DO DOCUMENTO (CÓDIGO 01, 1B, 04, 55 e
'REGISTRO C500: NOTA FISCAL/CONTA DE ENERGIA ELÉTRICA (CÓDIGO 06)
'REGISTRO C510: ITENS   DO   DOCUMENTO   NOTA   FISCAL/CONTA   ENERGIA ELÉTRICA
'REGISTRO C800: CUPOM FISCAL ELETRÔNICO – SAT (CF-E-SAT) (CÓDIGO 59)
'REGISTRO C990: ENCERRAMENTO DO BLOCO C


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



'C001
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


'NOTAS DE ENTRADA COMPRA
'cC100
Dim clin170 As Integer
Dim cID170 As Integer
'IMPORTANTE: para documentos de entrada, os campos de valor de imposto,
'base de cálculo e alíquota só devem ser informados se o adquirente tiver direito à apropriação do crédito
'(enfoque do declarante).

Do Until rsCompra.EOF
If rsCompra!IdFornecedor = 1131 Then
Else
    If rsCompra!chavenf Like "ALUGUEL*" Then
    Else
    cC100 = "|" & "C100" & "|" & "0" & "|" & "1" & "|" & rsCompra!IdFornecedor & "|" & "55" & "|" & "00" & "|" & rsCompra!Serie & "|" & rsCompra!NumNF & "|" & rsCompra!chavenf & "|" & Replace(rsCompra!DataEmissao, "/", "") & "|" & Replace(rsCompra!DataEmissao, "/", "") & "|" & Round(rsCompra!VlrTOTALNF, 2) & "|" & "0" & "|" & Round(rsCompra!VlrDesconto, 2) & "||" & Round(rsCompra!VlrTotalProdutos, 2) & "|" & "9" & "|" & "|" & "|" & "|" & Round(rsCompra!BaseCalculo, 2) & "|" & Round(rsCompra!Valor_ICMS, 2) & "|" & Round(rsCompra!BaseCalc_ST, 2) & "|" & Round(rsCompra!Valor_ICMS_ST, 2) & "|" & Round(rsCompra!Valor_IPI, 2) & "|" & Round(rsCompra!Valor_PIS, 2) & "|" & Round(rsCompra!Valor_Cofins, 2) & "|" & "|" & "|"
    
Print #iArq, cC100
cLinC = cLinC + 1
lC100 = lC100 + 1

    'cC170
    'COMPRAS
    cC170 = ""
    cID170 = rsCompra!ID
    clin170 = 1
    
    'Set rsCompraDet = DB.OpenRecordset("SELECT tbCompras.ID as ID, tbComprasDet.CST_ICMS, tbComprasDet.CST_IPI, tbComprasDet.CST_PIS, tbComprasDet.CST_Cofins, tbComprasDet.ID as ID_DET,tbComprasDet.IDCompra as ID_Compra,  tbComprasDet.IDProd, tbComprasDet.Qnt, tbCadProd.Unid, tbCompras.dataemissao, tbComprasDet.ValorTot, tbComprasDet.VlrDesc, tbComprasDet.CST, tbComprasDet.CFOP, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.BaseCalculo, tbComprasDet.Aliq_ICMS, tbComprasDet.Valor_ICMS, tbComprasDet.BaseCalc_ST, tbComprasDet.Aliq_ICMS_ST, tbComprasDet.Valor_ICMS_ST, tbComprasDet.Aliq_IPI, tbComprasDet.Valor_IPI, tbComprasDet.Aliq_PIS, tbComprasDet.Aliq_Cofins, tbComprasDet.Valor_PIS, tbComprasDet.Valor_Cofins " & _
    "FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
    "WHERE dataemissao >= #" & cDtINI & "# and dataemissao <= #" & cDtFIM & "# and tbComprasDet.IDCompra = " & cID170 & " ORDER BY tbCompras.ID, tbComprasDet.ID;")
    
    Set rsCompraDet = Db.OpenRecordset("SELECT tbCompras.ID AS ID, tbComprasDet.CST_ICMS, tbComprasDet.CST_IPI, tbComprasDet.CST_PIS, tbComprasDet.CST_Cofins, tbComprasDet.ID AS ID_DET, tbComprasDet.IDCompra AS ID_Compra, tbComprasDet.IDProd, tbComprasDet.Qnt, tbCadProd.Unid, tbCompras.dataemissao, tbComprasDet.ValorTot, tbComprasDet.VlrDesc, tbComprasDet.CST, tbComprasDet.CFOP, tbComprasDet.CFOP_ESCRITURADA, " & _
    "IIf(lancfiscal='CREDITO',tbComprasDet.BaseCalculo,0) as BaseCalculo, IIf(lancfiscal='CREDITO',tbComprasDet.Aliq_ICMS,0) AS Aliq_ICMS, IIf(lancfiscal='CREDITO',tbComprasDet.Valor_ICMS,0) AS Valor_ICMS, tbComprasDet.BaseCalc_ST, tbComprasDet.Aliq_ICMS_ST, tbComprasDet.Valor_ICMS_ST, IIf(lancfiscal='CREDITO',tbComprasDet.Aliq_IPI,0) AS Aliq_IPI, IIf(lancfiscal='CREDITO',tbComprasDet.Valor_IPI,0) AS Valor_IPI, IIf(lancfiscal='CREDITO',tbComprasDet.Aliq_PIS,0) AS Aliq_PIS, IIf(lancfiscal='CREDITO',tbComprasDet.Aliq_Cofins,0) AS Aliq_Cofins, IIf(lancfiscal='CREDITO',tbComprasDet.Valor_PIS,0) AS Valor_PIS, " & _
    "IIf(lancfiscal='CREDITO',tbComprasDet.Valor_Cofins,0) AS Valor_Cofins, tbComprasDet.LancFiscal " & _
    "FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
    "WHERE dataemissao >= #" & cDtIniVb & "# and dataemissao <= #" & cDtFimVb & "# and tbComprasDet.IDCompra = " & cID170 & " ORDER BY tbCompras.ID, tbComprasDet.ID;")
    'Somente informar impostos caso tenha direito a credito, senaõ imposto zero
        
    clin170 = 1
    Do Until rsCompraDet.EOF
    
    cC170 = "|" & "C170" & "|" & clin170 & "|" & rsCompraDet!IDProd & "||" & rsCompraDet!Qnt & "|" & rsCompraDet!Unid & "|" & Round(rsCompraDet!ValorTot, 2) & "|" & Round(rsCompraDet!VlrDesc, 2) & "|" & "0" & "|" & rsCompraDet!CST_ICMS & "|" & rsCompraDet!CFOP_ESCRITURADA & "|" & rsCompraDet!CFOP_ESCRITURADA & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_ICMS & "|" & Round(rsCompraDet!Valor_ICMS, 2) & "|" & Round(rsCompraDet!BaseCalc_ST, 2) & "|" & rsCompraDet!Aliq_ICMS_ST & "|" & Round(rsCompraDet!Valor_ICMS_ST, 2) & "|" & "0" & "|" & rsCompraDet!CST_IPI & "|" & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_IPI & "|" & Round(rsCompraDet!Valor_IPI, 2) & "|" & rsCompraDet!CST_PIS & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_PIS & "|||" & Round(rsCompraDet!Valor_PIS, 2) & "|" & rsCompraDet!CST_Cofins & "|" & Round(rsCompraDet!BaseCalculo, 2) & "|" & rsCompraDet!Aliq_Cofins & "|||" & Round(rsCompraDet!Valor_Cofins, 2) & "|" & "01" & "|" & "0" & "|"
    Print #iArq, cC170
    clin170 = clin170 + 1
    rsCompraDet.MoveNext
    cLinC = cLinC + 1
    lC170 = lC170 + 1
    Loop
    'cC170
    
    'c190
    
    'Set rsRegEntr = DB.OpenRecordset("SELECT Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbComprasDet.CST_ICMS, tbComprasDet.Aliq_ICMS, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC AS CFOP Desc, tbComprasDet.lancfiscal AS Lanc FIscal, Sum(tbComprasDet.ValorTot) AS Valor Contábil, Sum(tbComprasDet.BaseCalculo) AS Base de Calculo, Sum(tbComprasDet.Valor_ICMS) AS ICMS, Sum(tbComprasDet.Valor_IPI) AS IPI, Sum(tbComprasDet.Valor_PIS) AS PIS, Sum(tbComprasDet.Valor_Cofins) AS Cofins, Sum(tbComprasDet.Valor_ICMS_ST) AS ICMS_ST, Sum(tbComprasDet.BaseCalc_ST) AS BaseCalc_ST, tbCompras.ChaveNF, tbcompras.VlrTOTALNF " & _
    "FROM tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
    "WHERE (((tbCompras.DataEmissao) >= # " & cDtINI & "# And (tbCompras.DataEmissao) <= #" & cDtFIM & "#) And ((tbCadProd.IDProd) <> 2800)) " & _
    "GROUP BY Year(DataEmissao), Month(DataEmissao), tbComprasDet.CST_ICMS, tbComprasDet.Aliq_ICMS, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.lancfiscal, tbCompras.ChaveNF, tbcompras.VlrTOTALNF " & _
    "HAVING (((tbCompras.ChaveNF) = '" & rsCompra!ChaveNF & "')) " & _
    "ORDER BY Year(DataEmissao), Month(DataEmissao);")
    
    Set rsRegEntr = Db.OpenRecordset("SELECT q1.ANO, q1.MES, q1.CST_ICMS, q1.Aliq_ICMS, q1.CFOP, q1.[CFOP Desc], q1.[Lanc FIscal], Sum(q1.[Valor Contábil]) AS [Valor Contábil], Sum(q1.[Base de Calculo]) AS [Base de Calculo], Sum(q1.ICMS) AS ICMS, Sum(q1.IPI) AS IPI, Sum(q1.PIS) AS PIS, Sum(q1.Cofins) AS Cofins, Sum(q1.ICMS_ST) AS ICMS_ST, Sum(q1.BaseCalc_ST) AS BaseCalc_ST, Sum(q1.ChaveNF) AS ChaveNF, Sum(q1.VlrTOTALNF) AS VlrTOTALNF, q1.LancFiscal FROM ( " & _
    "SELECT Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbComprasDet.CST_ICMS, IIf(LancFiscal='CREDITO',tbComprasDet.Aliq_ICMS,0) AS Aliq_ICMS, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC AS [CFOP Desc], tbComprasDet.lancfiscal AS [Lanc FIscal], Sum(tbComprasDet.ValorTot) AS [Valor Contábil], Sum(IIf(LancFiscal='CREDITO',BaseCalculo,0)) AS [Base de Calculo], Sum(IIf(LancFiscal='CREDITO',Valor_ICMS,0)) AS ICMS, Sum(IIf(LancFiscal='CREDITO',Valor_IPI,0)) AS IPI, Sum(IIf(LancFiscal='CREDITO',Valor_PIS,0)) AS PIS, Sum(IIf(LancFiscal='CREDITO',Valor_Cofins,0)) AS Cofins, Sum(tbComprasDet.Valor_ICMS_ST) AS ICMS_ST, Sum(tbComprasDet.BaseCalc_ST) AS BaseCalc_ST, tbCompras.ChaveNF, tbCompras.VlrTOTALNF, tbComprasDet.LancFiscal " & _
    "FROM tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
    "WHERE (((tbCompras.DataEmissao) >= # " & cDtIniVb & "# And (tbCompras.DataEmissao) <= #" & cDtFimVb & "#) And ((tbCadProd.IDProd) <> 2800)) " & _
    "GROUP BY Year(DataEmissao), Month(DataEmissao), tbComprasDet.CST_ICMS, tbComprasDet.Aliq_ICMS, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.lancfiscal, tbCompras.ChaveNF, tbCompras.VlrTOTALNF, tbComprasDet.LancFiscal " & _
    "HAVING (((tbCompras.ChaveNF)='" & rsCompra!chavenf & "')) " & _
    "ORDER BY Year(DataEmissao), Month(DataEmissao) )  AS q1 " & _
    "GROUP BY q1.ANO, q1.MES, q1.CST_ICMS, q1.Aliq_ICMS, q1.CFOP, q1.[CFOP Desc], q1.[Lanc FIscal], q1.LancFiscal;")

    C190 = ""
    Do Until rsRegEntr.EOF
    cC190 = "|" & "C190" & "|" & rsRegEntr!CST_ICMS & "|" & rsRegEntr!CFOP & "|" & rsRegEntr!Aliq_ICMS & "|" & Round(rsRegEntr![Base de Calculo] + rsRegEntr!IPI, 2) & "|" & Round(rsRegEntr![Base de Calculo], 2) & "|" & Round(rsRegEntr!ICMS, 2) & "|" & Round(rsRegEntr!BaseCalc_ST, 2) & "|" & Round(rsRegEntr!ICMS_ST, 2) & "|" & "0" & "|" & Round(rsRegEntr!IPI, 2) & "|" & "|"
    Print #iArq, cC190
    rsRegEntr.MoveNext
    cLinC = cLinC + 1
    lC190 = lC190 + 1
    Loop

    rsRegEntr.Close
    
    rsCompraDet.Close
End If
End If
rsCompra.MoveNext
Loop

'NOTAS DE ENTRADA CIAP E OUTRAS ENTRADAS EMITIDAS POR MIM
'cC100
Do Until rsCIAP.EOF

cC100 = "|" & "C100" & "|" & "0" & "|" & "1" & "|" & rsCIAP!IdCliente & "|" & "55" & "|" & "00" & "|" & rsCIAP!Serie & "|" & rsCIAP!NumNF & "|" & rsCIAP!chavenf & "|" & Replace(rsCIAP!DataEmissao, "/", "") & "|" & Replace(rsCIAP!DataEmissao, "/", "") & "|" & Round(rsCIAP!VlrTOTALNF, 2) & "|" & "0" & "|" & Round(rsCIAP!VlrDesconto, 2) & "||" & Round(rsCIAP!VlrTotalProdutos, 2) & "|" & "9" & "|" & "|" & "|" & "|" & Round(rsCIAP!ICMS_BaseCalc, 2) & "|" & Round(rsCIAP!ICMS_Valor, 2) & "|" & Round(rsCIAP!ICMS_ST_BaseCalc, 2) & "|" & Round(rsCIAP!ICMS_ST_Valor, 2) & "|" & Round(rsCIAP!IPI_Valor, 2) & "|" & Round(rsCIAP!PIS_Valor, 2) & "|" & Round(rsCIAP!COFINS_Valor, 2) & "|" & "|" & "|"
    
Print #iArq, cC100
cLinC = cLinC + 1
lC100 = lC100 + 1

    'cC170
    'COMPRAS
    cC170 = ""
    cID170 = rsCIAP!ID
    clin170 = 1
    
    'Set rsCompraDet = DB.OpenRecordset("SELECT tbCompras.ID as ID, tbComprasDet.CST_ICMS, tbComprasDet.CST_IPI, tbComprasDet.CST_PIS, tbComprasDet.CST_Cofins, tbComprasDet.ID as ID_DET,tbComprasDet.IDCompra as ID_Compra,  tbComprasDet.IDProd, tbComprasDet.Qnt, tbCadProd.Unid, tbCompras.dataemissao, tbComprasDet.ValorTot, tbComprasDet.VlrDesc, tbComprasDet.CST, tbComprasDet.CFOP, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.BaseCalculo, tbComprasDet.Aliq_ICMS, tbComprasDet.Valor_ICMS, tbComprasDet.BaseCalc_ST, tbComprasDet.Aliq_ICMS_ST, tbComprasDet.Valor_ICMS_ST, tbComprasDet.Aliq_IPI, tbComprasDet.Valor_IPI, tbComprasDet.Aliq_PIS, tbComprasDet.Aliq_Cofins, tbComprasDet.Valor_PIS, tbComprasDet.Valor_Cofins " & _
    "FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
    "WHERE dataemissao >= #" & cDtINI & "# and dataemissao <= #" & cDtFIM & "# and tbComprasDet.IDCompra = " & cID170 & " ORDER BY tbCompras.ID, tbComprasDet.ID;")
    
    Set rsCIAPDet = Db.OpenRecordset("SELECT tbVendas.ID AS ID, tbVendasDet.CST_ICMS, tbVendasDet.CST_IPI, tbVendasDet.CST_PIS, tbVendasDet.CST_Cofins, tbVendasDet.ID AS ID_DET, tbVendasDet.IDVenda AS ID_Venda, tbVendasDet.IDProd, tbVendasDet.Qnt, tbCadProd.Unid, tbVendas.dataemissao, tbVendasDet.ValorTot, tbVendasDet.VlrDesc, tbVendasDet.CST, tbVendasDet.CFOP, tbVendasDet.CFOP_ESCRITURADA, " & _
    "tbVendasDet.BaseCalculo, tbVendasDet.Aliq_ICMS, tbVendasDet.Valor_ICMS, tbVendasDet.BaseCalc_ST, tbVendasDet.Aliq_ICMS_ST, tbVendasDet.Valor_ICMS_ST, tbVendasDet.Aliq_IPI, tbVendasDet.Valor_IPI, tbVendasDet.Aliq_PIS, tbVendasDet.Aliq_Cofins, tbVendasDet.Valor_PIS, " & _
    "tbVendasDet.Valor_Cofins, tbVendasDet.LancFiscal " & _
    "FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda " & _
    "WHERE dataemissao >= #" & cDtIni & "# and dataemissao <= #" & cDtFim & "# and tbVendasDet.IDVenda = " & cID170 & " and tbVendas.TipoNF = '0-ENTRADA' ORDER BY tbVendas.ID, tbVendasDet.ID;")
    
        
    clin170 = 1
    Do Until rsCIAPDet.EOF
    
    cC170 = "|" & "C170" & "|" & clin170 & "|" & rsCIAPDet!IDProd & "||" & rsCIAPDet!Qnt & "|" & rsCIAPDet!Unid & "|" & Round(rsCIAPDet!ValorTot, 2) & "|" & Round(rsCIAPDet!VlrDesc, 2) & "|" & "0" & "|" & rsCIAPDet!CST_ICMS & "|" & rsCIAPDet!CFOP_ESCRITURADA & "|" & rsCIAPDet!CFOP_ESCRITURADA & "|" & Round(rsCIAPDet!BaseCalculo, 2) & "|" & rsCIAPDet!Aliq_ICMS & "|" & Round(rsCIAPDet!Valor_ICMS, 2) & "|" & Round(rsCIAPDet!BaseCalc_ST, 2) & "|" & rsCIAPDet!Aliq_ICMS_ST & "|" & Round(rsCIAPDet!Valor_ICMS_ST, 2) & "|" & "0" & "|" & rsCIAPDet!CST_IPI & "|" & "|" & Round(rsCIAPDet!BaseCalculo, 2) & "|" & rsCIAPDet!Aliq_IPI & "|" & Round(rsCIAPDet!Valor_IPI, 2) & "|" & rsCIAPDet!CST_PIS & "|" & Round(rsCIAPDet!BaseCalculo, 2) & "|" & rsCIAPDet!Aliq_PIS & "|||" & Round(rsCIAPDet!Valor_PIS, 2) & "|" & rsCIAPDet!CST_Cofins & "|" & Round(rsCIAPDet!BaseCalculo, 2) & "|" & rsCIAPDet!Aliq_Cofins & "|||" & Round(rsCIAPDet!Valor_Cofins, 2) & "|" & "01" & "|" & "0" & "|"
    Print #iArq, cC170
    clin170 = clin170 + 1
    rsCIAPDet.MoveNext
    cLinC = cLinC + 1
    lC170 = lC170 + 1
    Loop
    'cC170
    
    'c190
    
    'Set rsRegEntr = DB.OpenRecordset("SELECT Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbComprasDet.CST_ICMS, tbComprasDet.Aliq_ICMS, tbComprasDet.CFOP_ESCRITURADA AS CFOP, tbComprasDet.CFOP_ESC_DESC AS CFOP Desc, tbComprasDet.lancfiscal AS Lanc FIscal, Sum(tbComprasDet.ValorTot) AS Valor Contábil, Sum(tbComprasDet.BaseCalculo) AS Base de Calculo, Sum(tbComprasDet.Valor_ICMS) AS ICMS, Sum(tbComprasDet.Valor_IPI) AS IPI, Sum(tbComprasDet.Valor_PIS) AS PIS, Sum(tbComprasDet.Valor_Cofins) AS Cofins, Sum(tbComprasDet.Valor_ICMS_ST) AS ICMS_ST, Sum(tbComprasDet.BaseCalc_ST) AS BaseCalc_ST, tbCompras.ChaveNF, tbcompras.VlrTOTALNF " & _
    "FROM tbFornecedor INNER JOIN (tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra) ON tbFornecedor.IDFor = tbCompras.IdFornecedor " & _
    "WHERE (((tbCompras.DataEmissao) >= # " & cDtINI & "# And (tbCompras.DataEmissao) <= #" & cDtFIM & "#) And ((tbCadProd.IDProd) <> 2800)) " & _
    "GROUP BY Year(DataEmissao), Month(DataEmissao), tbComprasDet.CST_ICMS, tbComprasDet.Aliq_ICMS, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC, tbComprasDet.lancfiscal, tbCompras.ChaveNF, tbcompras.VlrTOTALNF " & _
    "HAVING (((tbCompras.ChaveNF) = '" & rsCompra!ChaveNF & "')) " & _
    "ORDER BY Year(DataEmissao), Month(DataEmissao);")
    
    Set rsRegEntr = Db.OpenRecordset("SELECT Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbVendasDet.CST_ICMS, tbVendasDet.Aliq_ICMS, tbVendasDet.CFOP_ESCRITURADA AS CFOP, tbVendasDet.CFOP_ESC_DESC AS [CFOP Desc], tbVendasDet.lancfiscal AS [Lanc FIscal], Sum(tbVendasDet.ValorTot) AS [Valor Contábil], sum(ICMS_BaseCalc) AS [Base de Calculo], sum(ICMS_Valor) AS ICMS, sum(IPI_Valor) AS IPI, sum(PIS_Valor) AS PIS, sum(Cofins_Valor) AS Cofins, Sum(tbVendasDet.Valor_ICMS_ST) AS ICMS_ST, Sum(tbVendasDet.BaseCalc_ST) AS BaseCalc_ST, tbVendas.ChaveNF, tbVendas.VlrTOTALNF, tbVendasDet.LancFiscal " & _
    "FROM tbCliente INNER JOIN (tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda) ON tbCliente.IDCliente = tbVendas.IdCliente " & _
    "WHERE (((tbVendas.DataEmissao) >= # " & cDtIni & "# And (tbVendas.DataEmissao) <= #" & cDtFim & "#) And ((tbCadProd.IDProd) <> 2800)) and tbVendas.tipoNF = '0-ENTRADA' " & _
    "GROUP BY Year(DataEmissao), Month(DataEmissao), tbVendasDet.CST_ICMS, tbVendasDet.Aliq_ICMS, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, tbVendasDet.lancfiscal, tbVendas.ChaveNF, tbVendas.VlrTOTALNF, tbVendasDet.LancFiscal " & _
    "HAVING (((tbVendas.ChaveNF)='" & rsCIAP!chavenf & "')) " & _
    "ORDER BY Year(DataEmissao), Month(DataEmissao);")



    C190 = ""
    Do Until rsRegEntr.EOF
    cC190 = "|" & "C190" & "|" & rsRegEntr!CST_ICMS & "|" & rsRegEntr!CFOP & "|" & rsRegEntr!Aliq_ICMS & "|" & Round(rsRegEntr![Base de Calculo] + rsRegEntr!IPI, 2) & "|" & Round(rsRegEntr![Base de Calculo], 2) & "|" & Round(rsRegEntr!ICMS, 2) & "|" & Round(rsRegEntr!BaseCalc_ST, 2) & "|" & Round(rsRegEntr!ICMS_ST, 2) & "|" & "0" & "|" & Round(rsRegEntr!IPI, 2) & "|" & "|"
    Print #iArq, cC190
    rsRegEntr.MoveNext
    cLinC = cLinC + 1
    lC190 = lC190 + 1
    Loop

    rsRegEntr.Close
    
    rsCIAPDet.Close


rsCIAP.MoveNext
Loop


'NOTAS DE SAÍDA VENDA
'cC100
Do Until rsVenda.EOF
cC100 = "|" & "C100" & "|" & "1" & "|" & "0" & "|" & rsVenda!IdCliente & "|" & "55" & "|" & "00" & "|" & rsVenda!Serie & "|" & rsVenda!NumNF & "|" & rsVenda!chavenf & "|" & Replace(rsVenda!DataEmissao, "/", "") & "|" & Replace(rsVenda!DataEmissao, "/", "") & "|" & Round(rsVenda!VlrTOTALNF, 2) & "|" & "0" & "|" & Round(rsVenda!VlrDesconto, 2) & "||" & Round(rsVenda!VlrTotalProdutos, 2) & "|" & "9" & "|" & "|" & "|" & "|" & Round(rsVenda!ICMS_BaseCalc, 2) & "|" & Round(rsVenda!ICMS_Valor, 2) & "|" & Round(rsVenda!ICMS_ST_BaseCalc, 2) & "|" & Round(rsVenda!ICMS_ST_Valor, 2) & "|" & Round(rsVenda!IPI_Valor, 2) & "|" & Round(rsVenda!PIS_Valor, 2) & "|" & Round(rsVenda!COFINS_Valor, 2) & "|" & "|" & "|"
Print #iArq, cC100
cLinC = cLinC + 1
lC100 = lC100 + 1

'VENDAS
    'cC170

    cC170 = ""
    cID170 = rsVenda!ID
    'clin170 = 0
        
    Set rsVendaDet = Db.OpenRecordset("SELECT tbVendas.ID as ID, tbVendasDet.ID as ID_DET,tbVendasDet.IDVenda as ID_Venda,  tbVendasDet.IDProd, tbVendasDet.Qnt, tbCadProd.Unid, tbVendas.dataemissao, tbVendasDet.ValorTot, tbVendasDet.VlrDesc, tbVendasDet.CST, tbVendasDet.CFOP, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.BaseCalculo, tbVendasDet.Aliq_ICMS, tbVendasDet.Valor_ICMS, tbVendasDet.BaseCalc_ST, tbVendasDet.Aliq_ICMS_ST, tbVendasDet.Valor_ICMS_ST, tbVendasDet.Aliq_IPI, tbVendasDet.Valor_IPI, tbVendasDet.Aliq_PIS, tbVendasDet.Aliq_Cofins, tbVendasDet.Valor_PIS, tbVendasDet.Valor_Cofins, tbVendasDet.CST_ICMS, CST_IPI, CST_PIS, CST_Cofins " & _
    "FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda " & _
    "WHERE dataemissao >= #" & cDtIni & "# and dataemissao <= #" & cDtFim & "# and tbVendasDet.IDVenda = " & cID170 & " ORDER BY tbVendas.ID, tbVendasDet.ID;")

    
    Do Until rsVendaDet.EOF
    
    
    'cC170 = "|" & "C170" & "|" & clin170 & "|" & rsVendaDet!IDProd & "||" & rsVendaDet!Qnt & "|" & rsVendaDet!Unid & "|" & Round(rsVendaDet!ValorTot, 2) & "|" & Round(rsVendaDet!VlrDesc, 2) & "|" & "0" & "|" & rsVendaDet!CST_ICMS & "|" & rsVendaDet!CFOP_ESCRITURADA & "|" & rsVendaDet!CFOP_ESCRITURADA & "|" & Round(rsVendaDet!BaseCalculo, 2) & "|" & rsVendaDet!Aliq_ICMS & "|" & Round(rsVendaDet!Valor_ICMS, 2) & "|" & Round(rsVendaDet!BaseCalc_ST, 2) & "|" & rsVendaDet!Aliq_ICMS_ST & "|" & Round(rsVendaDet!Valor_ICMS_ST, 2) & "|" & "0" & "|" & rsVendaDet!CST_IPI & "|" & "|" & Round(rsVendaDet!BaseCalculo, 2) & "|" & rsVendaDet!Aliq_IPI & "|" & Round(rsVendaDet!Valor_IPI, 2) & "|" & rsVendaDet!CST_PIS & "|" & Round(rsVendaDet!BaseCalculo, 2) & "|" & rsVendaDet!Aliq_PIS & "|||" & Round(rsVendaDet!Valor_PIS, 2) & "|" & rsVendaDet!CST_Cofins & "|" & Round(rsVendaDet!BaseCalculo, 2) & "|" & rsVendaDet!Aliq_COFINS & "|||" & Round(rsVendaDet!Valor_COFINS, 2) & "|" & "01" & "|"
    
    
   'O PROGRAMA NÃO QUER O DETALHE NA VENDA, APENAS O ANAILTICO REGISTRO C190
   ' Print #iArq, cC170
   ' clin170 = clin170 + 1
    rsVendaDet.MoveNext
   ' cLinC = cLinC + 1
   ' lC170 = lC170 + 1
    Loop
    
    'c190
    
    Set rsRegSaida = Db.OpenRecordset("SELECT Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbVendasDet.CST_ICMS, tbVendasDet.Aliq_ICMS, tbVendasDet.CFOP_ESCRITURADA AS CFOP, tbVendasDet.CFOP_ESC_DESC AS [CFOP Desc], tbVendasDet.lancfiscal AS [Lanc FIscal], Sum(tbVendasDet.ValorTot) AS [Valor Contábil], Sum(tbVendasDet.BaseCalculo) AS [Base de Calculo], Sum(tbVendasDet.Valor_ICMS) AS ICMS, Sum(tbVendasDet.Valor_IPI) AS IPI, Sum(tbVendasDet.Valor_PIS) AS PIS, Sum(tbVendasDet.Valor_Cofins) AS Cofins, Sum(tbVendasDet.Valor_ICMS_ST) AS ICMS_ST, Sum(tbVendasDet.BaseCalc_ST) AS BaseCalc_ST, tbVendas.ChaveNF, tbVendas.VlrTOTALNF " & _
    "FROM (tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente) INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
    "WHERE (((tbVendas.DataEmissao) >= #" & cDtIni & "# And (tbVendas.DataEmissao) <= #" & cDtFim & "#)) " & _
    "GROUP BY Year(DataEmissao), Month(DataEmissao), tbVendasDet.CST_ICMS, tbVendasDet.Aliq_ICMS, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, tbVendasDet.lancfiscal, tbVendas.ChaveNF, tbVendas.VlrTOTALNF " & _
    "HAVING (((tbVendas.ChaveNF)='" & rsVenda!chavenf & "')) " & _
    "ORDER BY Year(DataEmissao), Month(DataEmissao);")


    C190 = ""
    Do Until rsRegSaida.EOF
    cC190 = "|" & "C190" & "|" & rsRegSaida!CST_ICMS & "|" & rsRegSaida!CFOP & "|" & rsRegSaida!Aliq_ICMS & "|" & Round(rsRegSaida![Base de Calculo] + rsRegSaida!IPI, 2) & "|" & Round(rsRegSaida![Base de Calculo], 2) & "|" & Round(rsRegSaida!ICMS, 2) & "|" & Round(rsRegSaida!BaseCalc_ST, 2) & "|" & Round(rsRegSaida!ICMS_ST, 2) & "|" & "0" & "|" & Round(rsRegSaida!IPI, 2) & "|" & "|"
    Print #iArq, cC190
    rsRegSaida.MoveNext
    cLinC = cLinC + 1
    lC190 = lC190 + 1
    Loop

    rsRegSaida.Close
    rsVendaDet.Close

rsVenda.MoveNext
Loop

'c101 Omitido
'c105 Omitido
'c110 Omitido


'C190
'C190 = ""

'Do Until rsRegEntr.EOF

'cC190 = "|" & "C190" & "|" & rsRegEntr!CST_ICMS & "|" & rsRegEntr!CFOP & "|" & rsRegEntr!Aliq_ICMS & "|" & Round(rsRegEntr!Valor Contábil, 2) & "|" & Round(rsRegEntr!Base de Calculo, 2) & "|" & Round(rsRegEntr!ICMS, 2) & "|" & Round(rsRegEntr!BaseCalc_ST, 2) & "|" & Round(rsRegEntr!ICMS_ST, 2) & "|" & "0" & "|" & Round(rsRegEntr!IPI, 2) & "|" & "|"
'Print #iArq, cC190
'rsRegEntr.MoveNext
'cLinC = cLinC + 1
'lC190 = lC190 + 1
'Loop

'Do Until rsRegSaida.EOF

'cC190 = "|" & "C190" & "|" & rsRegSaida!CST & "|" & rsRegSaida!CFOP & "|" & rsRegSaida!Aliq_ICMS & "|" & rsRegSaida!Valor Contábil & "|" & rsRegSaida!Base de Calculo & "|" & rsRegSaida!ICMS & "|" & rsRegSaida!BaseCalc_ST & "|" & rsRegSaida!ICMS_ST & "|" & "0" & "|" & rsRegSaida!IPI & "|" & "|"
'Print #iArq, cC190
'rsRegSaida.MoveNext
'cLinC = cLinC + 1
'lC190 = lC190 + 1
'Loop


'C500
cC500 = ""
rsEnergiaInjetada.MoveFirst

Do Until rsEnergia.EOF

'cC500 = "|" & "C500" & "|" & "0" & "|" & "1" & "|" & rsEnergia!IdFornecedor & "|" & "06" & "|" & "00" & "|" & rsEnergia!Serie & "||" & "04" & "|" & rsEnergia!NumNF & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & rsEnergia!VlrTotalProdutos & "|" & rsEnergia!VlrDesconto & "|" & rsEnergia!VlrTotalProdutos & "|" & "0" & "|" & "|" & "|" & rsEnergia!ICMS_BaseCalc & "|" & rsEnergia!ICMS_Valor & "|" & "0" & "|" & "0" & "|" & "|" & rsEnergia!PIS_Valor & "|" & rsEnergia!COFINS_Valor & "|" & "2" & "|" & "12" & "|" & "|" & "|" & "|" & "1" & "|" & "3552205" & "|" & "3.3.1.05" & "|"
'cC500 = "|" & "C500" & "|" & "0" & "|" & "1" & "|" & rsEnergia!IdFornecedor & "|" & "06" & "|" & "00" & "|" & rsEnergia!Serie & "||" & "04" & "|" & rsEnergia!NumNF & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & rsEnergia!VlrTotalProdutos & "|" & rsEnergia!VlrDesconto & "|" & rsEnergia!VlrTotalProdutos & "|" & "0" & "|" & "|" & "|" & rsEnergia!ICMS_BaseCalc & "|" & rsEnergia!ICMS_Valor & "|" & "0" & "|" & "0" & "|" & "|" & rsEnergia!PIS_Valor & "|" & rsEnergia!COFINS_Valor & "|" & "2" & "|" & "12" & "|" & "|" & "|" & "|" & "1" & "|" & "3552205" & "|" & "3.3.1.05" & "|" & "06" & "|" & rsEnergiaInjetada!Hash_Conta_Ref & "|" & rsEnergiaInjetada!Serie_Ref & "|" & rsEnergiaInjetada!Num_nota_Ref & "|" & rsEnergiaInjetada!Mes_Ref & rsEnergiaInjetada!Ano_Ref & "|" & rsEnergiaInjetada!Energia_Injetada_kWh & "|" & rsEnergiaInjetada!Outras_Deducoes & "|"
cC500 = "|" & "C500" & "|" & "0" & "|" & "1" & "|" & rsEnergia!IdFornecedor & "|" & "06" & "|" & "00" & "|" & rsEnergia!Serie & "||" & "04" & "|" & rsEnergia!NumNF & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & Replace(rsEnergia!DataEmissao, "/", "") & "|" & rsEnergia!VlrTotalProdutos & "|" & rsEnergia!VlrDesconto & "|" & rsEnergia!VlrTotalProdutos & "|" & "0" & "|" & "|" & "|" & rsEnergia!ICMS_BaseCalc & "|" & rsEnergia!ICMS_Valor & "|" & "0" & "|" & "0" & "|" & "|" & rsEnergia!PIS_Valor & "|" & rsEnergia!COFINS_Valor & "|" & "2" & "|" & "12" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|"
Print #iArq, cC500
rsEnergia.MoveNext
cLinC = cLinC + 1
lC500 = lC500 + 1
Loop


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
Print #iArq, cC590
rsEnergiaDet.MoveNext
cLinC = cLinC + 1
lC590 = lC590 + 1
Loop

'C800
cC800 = ""
Set rsSAT = Db.OpenRecordset("select * from tbVendasSAT WHERE tbVendasSAT.DataEmissao  >= # " & cDtIni & " 00:00:00" & " # And tbVendasSAT.DataEmissao <= # " & cDtFim & " 23:59:59" & " #;")
Do Until rsSAT.EOF
cC800 = "|" & "C800" & "|" & "59" & "|" & "00" & "|" & rsSAT!numCF & "|" & Replace(Format(rsSAT!DataEmissao, "dd/mm/yyyy"), "/", "") & "|" & rsSAT!Vlr_TotalCF & "|" & rsSAT!vPIS & "|" & rsSAT!vCofins & "|" & rsSAT!CPF_CNPJ & "|" & rsSAT!NumSerieSAT & "|" & rsSAT!ChaveCF & "|" & rsSAT!vDesc & "|" & rsSAT!Vlr_TotalCF & "|" & "0" & "|" & rsSAT!vICMS & "|" & "0" & "|" & "0" & "|"
Print #iArq, cC800
cLinC = cLinC + 1
lC800 = lC800 + 1

    'C810
    cC810 = ""
    Set rsSATDet = Db.OpenRecordset("select * from tbVendasSATDet where idSAT = " & rsSAT!idSAT & "")
    Do Until rsSATDet.EOF
    cC810 = "|" & "C810" & "|" & rsSATDet!NumItem & "|" & rsSATDet!IDProd & "|" & rsSATDet!Qt & "|" & rsSATDet!UN_Com & "|" & rsSATDet!Vlr_Item & "|" & rsSATDet!CST_ICMS & "|" & rsSATDet!CFOP & "|"
    Print #iArq, cC810
    rsSATDet.MoveNext
    cLinC = cLinC + 1
    lC810 = lC810 + 1
    Loop
    
            'C850
            cC850 = ""
            Set rsSATResumo = Db.OpenRecordset("SELECT tbVendasDet.CST_ICMS, tbVendasDet.CFOP, tbVendasDet.Aliq_ICMS AS Aliq_ICMS, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo,SUM(tbVendasDet.ValorTot) as ValorTot, Sum(tbVendasDet.Valor_ICMS) AS Valor_ICMS " & _
            "FROM (tbVendasDet INNER JOIN tbVendas ON tbVendasDet.IDVenda = tbVendas.ID) INNER JOIN tbVendasSAT ON tbVendas.ChaveNF = tbVendasSAT.ChaveCF " & _
            "WHERE tbVendas.ChaveNF = '" & rsSAT!ChaveCF & "' " & _
            "GROUP BY tbVendasDet.CST_ICMS, tbVendasDet.CFOP, tbVendasDet.Aliq_ICMS;")


            Do Until rsSATResumo.EOF
            cC850 = "|" & "C850" & "|" & rsSATResumo!CST_ICMS & "|" & rsSATResumo!CFOP & "|" & rsSATResumo!Aliq_ICMS & "|" & rsSATResumo!ValorTot & "|" & rsSATResumo!BaseCalculo & "|" & rsSATResumo!Valor_ICMS & "|" & "|"
            Print #iArq, cC850
            rsSATResumo.MoveNext
            cLinC = cLinC + 1
            lC850 = lC850 + 1
            Loop


rsSAT.MoveNext
Loop



SemDadosC001:

If cC100 = "" Then
cDtC100 = "NAO"
Else
End If

'C990
cLinC = cLinC + 1
lC990 = lC990 + 1
cC990 = "|" & "C990" & "|" & cLinC & "|"
Print #iArq, cC990




'D - Documentos Fiscais II – Serviços (ICMS)


Dim rsTransp As DAO.Recordset
'Set rsTransp = "SELECT tbTransportes.*, tbMunicipioIBGE.COD_MUNICIPIO AS Rem_IBGE, tbMunicipioIBGE_1.COD_MUNICIPIO AS Dest_IBGE,  tbTransportes.DataEmissao " & _
'"FROM (tbTransportes INNER JOIN tbMunicipioIBGE ON tbTransportes.RemetenteCidade = tbMunicipioIBGE.MUNICIPIO_ACENTO) INNER JOIN tbMunicipioIBGE AS tbMunicipioIBGE_1 ON tbTransportes.DestinatarioCidade = tbMunicipioIBGE_1.MUNICIPIO_ACENTO " & _
'"WHERE (((tbTransportes.DataEmissao)>=#" & cDtINI & "# And (tbTransportes.DataEmissao)<=#" & cDtFIM & "#));"


Set rsTransp = Db.OpenRecordset("SELECT tbTransportes.*, tbMunicipioIBGE.COD_MUNICIPIO AS Rem_IBGE, tbMunicipioIBGE_1.COD_MUNICIPIO AS Dest_IBGE " & _
"FROM (tbTransportes LEFT OUTER JOIN tbMunicipioIBGE ON tbTransportes.RemetenteCidade = tbMunicipioIBGE.MUNICIPIO_ACENTO) LEFT OUTER JOIN tbMunicipioIBGE AS tbMunicipioIBGE_1 ON tbTransportes.DestinatarioCidade = tbMunicipioIBGE_1.MUNICIPIO_ACENTO " & _
"WHERE (((tbTransportes.DataEmissao)>=#" & cDtIni & "# And (tbTransportes.DataEmissao)<=#" & cDtFim & "#) and LancFiscal = 'CREDITO');")


Dim rsTrTotal As DAO.Recordset



Dim cLinBlocoD As Integer
cLinBlocoD = 0
'REGISTRO D001: ABERTURA DO BLOCO D
'REGISTRO D100: NOTA FISCAL DE SERVIÇO DE TRANSPORTE
'REGISTRO D990: ENCERRAMENTO DO BLOCO D.

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
cD001 = "|" & "D001" & "|" & "0" & "|"
cLinBlocoD = cLinBlocoD + 1
lD001 = lD001 + 1
Print #iArq, cD001
End If


'cD100

Do Until rsTransp.EOF
   
    'cD100 = "|" & "D100" & "|" & "0" & "|" & "1" & "|" & rsTransp!ID_Emit & "|" & "57" & "|" & "00" & "|" & rsTransp!Serie & "|" & "|" & rsTransp!Num_CTE & "|" & rsTransp!ChaveCTE & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & "|" & rsTransp!ChaveCTE & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & "0" & "|" & "1" & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & Round(rsTransp!BaseCalcICMS, 2) & "|" & Round(rsTransp!ValorICMS, 2) & "|" & "0" & "|" & "|" & "2" & "|" & rsTransp!REM_IBGE & "|" & rsTransp!DEST_IBGE & "|"
    cD100 = "|" & "D100" & "|" & "0" & "|" & "1" & "|" & rsTransp!ID_Emit & "|" & "57" & "|" & "00" & "|" & rsTransp!Serie & "|" & "|" & rsTransp!Num_CTE & "|" & rsTransp!ChaveCTE & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & Replace(rsTransp!DataEmissao, "/", "") & "|" & "|" & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & "0" & "|" & "1" & "|" & Round(rsTransp!ValorTotalServico, 2) & "|" & Round(rsTransp!BaseCalcICMS, 2) & "|" & Round(rsTransp!ValorICMS, 2) & "|" & "0" & "|" & "|" & "2" & "|" & rsTransp!REM_IBGE & "|" & rsTransp!DEST_IBGE & "|"
   
Print #iArq, cD100
cLinBlocoD = cLinBlocoD + 1
lD100 = lD100 + 1

    'cD190
    Set rsTrTotal = Db.OpenRecordset("SELECT tbTransportes.CST, tbTransportes.CFOP_ESCRITURADA, tbTransportes.AliqICMS, Sum(tbTransportes.ValorTotalServico) AS SomaDeValorTotalServico, Sum(tbTransportes.BaseCalcICMS) AS SomaDeBaseCalcICMS, Sum(tbTransportes.ValorICMS) AS SomaDeValorICMS " & _
    "FROM tbTransportes " & _
    "WHERE (((tbTransportes.DataEmissao) >= #" & cDtIni & "# And (tbTransportes.DataEmissao) <= #" & cDtFim & "#)) " & _
    "GROUP BY tbTransportes.CST, tbTransportes.CFOP_ESCRITURADA, tbTransportes.AliqICMS,tbTransportes.ChaveCTE " & _
    "HAVING tbTransportes.ChaveCTE='" & rsTransp!ChaveCTE & "';")

    Do Until rsTrTotal.EOF
    cD190 = "|" & "D190" & "|" & "0" & rsTrTotal!CST & "|" & rsTrTotal!CFOP_ESCRITURADA & "|" & rsTrTotal!AliqICMS & "|" & Round(rsTrTotal!SomaDeValorTotalServico, 2) & "|" & Round(rsTrTotal!SomaDeBaseCalcICMS, 2) & "|" & Round(rsTrTotal!SomadeValorICMS, 2) & "|" & "0" & "|" & "|"
    Print #iArq, cD190
    cLinBlocoD = cLinBlocoD + 1
    lD190 = lD190 + 1
    rsTrTotal.MoveNext
    Loop
    
    rsTrTotal.Close

rsTransp.MoveNext

Loop


semdatatransportes:

'cD990
cLinBlocoD = cLinBlocoD + 1
lD990 = lD990 + 1
cD990 = "|" & "D990" & "|" & cLinBlocoD & "|"
Print #iArq, cD990



'E - Apuração do ICMS e do IPI
'REGISTRO E001: ABERTURA DO BLOCO E
'REGISTRO E100: PERÍODO DA APURAÇÃO DO ICMS
'REGISTRO E110: APURAÇÃO DO ICMS – OPERAÇÕES PRÓPRIAS
'REGISTRO E200: PERÍODO DA APURAÇÃO DO ICMS - SUBSTITUIÇÃO TRIBUTÁRIA.
'REGISTRO E210: APURAÇÃO DO ICMS – SUBSTITUIÇÃO TRIBUTÁRIA.
'REGISTRO E500: PERÍODO DE APURAÇÃO DO IPI.
'REGISTRO E510: CONSOLIDAÇÃO DOS VALORES DO IPI.
'REGISTRO E520: APURAÇÃO DO IPI.


Dim rsResumoICMS As DAO.Recordset
'Set rsResumoICMS = DB.OpenRecordset("SELECT * FROM tbResumo_ICMS where ano = year('" & cDtINI & "') and mes = '" & Int(Left(cDtINI, 2)) & "'")
'Set rsResumoICMS = Db.OpenRecordset("SELECT tbResumo_ICMS.ANO, tbResumo_ICMS.MES, tbResumo_ICMS.CRED, tbResumo_ICMS.DEB, tbResumo_ICMS.Saldo_Mes, tbResumo_ICMS.CRED_MES_ANT, tbResumo_ICMS.CRED_TRANSPORTAR, tbResumo_ICMS.SALDO, cstAcumCiap.Valor_ICMS AS CIAP " & _
'"FROM tbResumo_ICMS INNER JOIN cstAcumCiap ON (tbResumo_ICMS.MES = cstAcumCiap.MES) AND (tbResumo_ICMS.ANO = cstAcumCiap.ANO) " & _
'"WHERE (((tbResumo_ICMS.ANO)=Year('" & cDtIniVb & "')) AND ((tbResumo_ICMS.MES)=" & Int(Left(cDtIniVb, 2)) & "));")

Set rsResumoICMS = Db.OpenRecordset("SELECT tbResumo_ICMS.ANO, tbResumo_ICMS.MES, tbResumo_ICMS.CRED, tbResumo_ICMS.DEB, tbResumo_ICMS.CIAP_EM_NF, tbResumo_ICMS.Saldo_Mes, tbResumo_ICMS.CRED_MES_ANT, tbResumo_ICMS.CRED_TRANSPORTAR, tbResumo_ICMS.SALDO " & _
"FROM tbResumo_ICMS " & _
"WHERE (((tbResumo_ICMS.ANO)=Year('" & cDtIniVb & "')) AND ((tbResumo_ICMS.MES)=" & Int(Left(cDtIniVb, 2)) & "));")


Dim cLinE As Integer
cLinE = 0

'E001
cE001 = "|" & "E001" & "|" & "0" & "|"
Print #iArq, cE001
cLinE = cLinE + 1
lE001 = lE001 + 1

'E100
cE100 = "|" & "E100" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|"
Print #iArq, cE100
cLinE = cLinE + 1
lE100 = lE100 + 1

'E110
If rsResumoICMS!SALDO <= 0 Then
cSaldoApurado = 0
cIcmsaRecolher = rsResumoICMS!SALDO * -1
Else
cSaldoApurado = rsResumoICMS!SALDO
cIcmsaRecolher = 0
End If

'- rsResumoICMS!CIAP tirei pra fazer a conta certa
If rsResumoICMS!CRED_TRANSPORTAR < 0 Then
cE110 = "|" & "E110" & "|" & Round(rsResumoICMS!DEB, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & Round(rsResumoICMS!CRED, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & Round(rsResumoICMS!CRED_MES_ANT, 2) & "|" & Round(cIcmsaRecolher, 2) & "|" & "0" & "|" & Round(cIcmsaRecolher, 2) & "|" & Round(rsResumoICMS!CRED_TRANSPORTAR, 2) * -1 & "|" & "0" & "|"
Else
cE110 = "|" & "E110" & "|" & Round(rsResumoICMS!DEB, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & Round(rsResumoICMS!CRED, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & Round(rsResumoICMS!CRED_MES_ANT, 2) & "|" & Round(cIcmsaRecolher, 2) & "|" & "0" & "|" & Round(cIcmsaRecolher, 2) & "|" & Round(rsResumoICMS!CRED_TRANSPORTAR, 2) & "|" & "0" & "|"
End If
Print #iArq, cE110
cLinE = cLinE + 1
lE110 = lE110 + 1


'cE116
Dim cAnomes As String

Select Case Int(Left(cDtIniVb, 2))
Case Is < 10
cAnomes = CStr("0" & Int(Left(cDtIniVb, 2)) & year(cDtIniVb))
Case Else
cAnomes = CStr(Int(Left(cDtIniVb, 2)) & year(cDtIniVb))
End Select

Select Case month(cDtFimVb)
Case Is < 10
cMes = CStr("0" & month(cDtFimVb))
Case Else
cMes = month(cDtFimVb)
End Select

cE116 = "|" & "E116" & "|" & "000" & "|" & Round(cIcmsaRecolher, 2) & "|" & "20" & cMes & "" & year(cDtFimVb) & "|" & "046-2" & "|" & "|" & "|" & "|" & "|" & cAnomes & "|"
Print #iArq, cE116
cLinE = cLinE + 1
lE116 = lE116 + 1

'cE200
'REGISTROS DE ICMS ST - incompleto
Dim rsValST As DAO.Recordset
Set rsValST = Db.OpenRecordset("SELECT Sum(tbComprasDet.Valor_ICMS_ST) AS ValorST " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= # " & cDtIniVb & " # And (tbCompras.DataEmissao) <= # " & cDtFimVb & " #)) ")

If rsValST.EOF = True And rsValST.BOF = True Then
Else
'cE200 = "|" & "E200" & "|" & "SP" & cSTR_DtINI & "|" & cSTR_DtFIM & "|"
'Print #iArq, cE200
'cLinE = cLinE + 1
'lE200 = lE200 + 1

 '   cE210 = "|"

End If

'cE210

'cE500
cE500 = "|" & "E500" & "|" & "0" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|"
Print #iArq, cE500
cLinE = cLinE + 1
lE500 = lE500 + 1

'cE510
'REGISTROS DE IPI DE COMPRA
Dim rsValIPI As DAO.Recordset
Set rsValIPI = Db.OpenRecordset("SELECT tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CST_IPI, Sum(tbComprasDet.ValorTot) AS ValorTot, Sum(tbComprasDet.BaseCalculo) AS BaseCalculo, Sum(tbComprasDet.Valor_IPI) AS Valor_IPI " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbCompras.DataEmissao) >= # " & cDtIniVb & " # And (tbCompras.DataEmissao) <= # " & cDtFimVb & " #) and tbComprasDet.LancFiscal='CREDITO') " & _
"GROUP BY tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CST_IPI;")

Do Until rsValIPI.EOF

cE510 = "|" & "E510" & "|" & rsValIPI!CFOP_ESCRITURADA & "|" & rsValIPI!CST_IPI & "|" & Round(rsValIPI!ValorTot, 2) & "|" & Round(rsValIPI!BaseCalculo, 2) & "|" & Round(rsValIPI!Valor_IPI, 2) & "|"
Print #iArq, cE510
cLinE = cLinE + 1
lE510 = lE510 + 1
rsValIPI.MoveNext
Loop
'REGISTROS DE IPI DE VENDAS
Set rsValIPI = Db.OpenRecordset("SELECT tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CST_IPI, Sum(tbVendasDet.ValorTot) AS ValorTot, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo, Sum(tbVendasDet.Valor_IPI) AS Valor_IPI " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE tbVendas.DataEmissao >= # " & cDtIniVb & " # And tbVendas.DataEmissao <= # " & cDtFimVb & " # and [status]='ATIVO' " & _
"GROUP BY tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CST_IPI;")

Do Until rsValIPI.EOF

cE510 = "|" & "E510" & "|" & rsValIPI!CFOP_ESCRITURADA & "|" & rsValIPI!CST_IPI & "|" & Round(rsValIPI!ValorTot, 2) & "|" & Round(rsValIPI!BaseCalculo, 2) & "|" & Round(rsValIPI!Valor_IPI, 2) & "|"
Print #iArq, cE510
cLinE = cLinE + 1
lE510 = lE510 + 1
rsValIPI.MoveNext
Loop


'cE520
Dim rsApuIPI As DAO.Recordset
Set rsApuIPI = Db.OpenRecordset("SELECT * FROM tbResumo_IPI " & _
"WHERE tbResumo_IPI.ANO='" & year(cDtIniVb) & "' AND tbResumo_IPI.MES=" & Int(Left(cDtIniVb, 2)) & ";")

Dim cSaldoIPI As Double
If rsApuIPI!SALDO >= 0 Then
cSaldoIPI = 0
Else
cSaldoIPI = rsApuIPI!SALDO * -1
End If
cE520 = "|" & "E520" & "|" & Round(rsApuIPI!CRED_MES_ANT, 2) & "|" & Round(rsApuIPI!DEB, 2) & "|" & Round(rsApuIPI!CRED, 2) & "|" & "0" & "|" & "0" & "|" & Round(rsApuIPI!CRED_TRANSPORTAR, 2) & "|" & Round(cSaldoIPI, 2) & "|"
Print #iArq, cE520
cLinE = cLinE + 1
lE520 = lE520 + 1

'cE990
cLinE = cLinE + 1
lE990 = lE990 + 1
cE990 = "|" & "E990" & "|" & cLinE & "|"
Print #iArq, cE990


'G* - Controle do Crédito de ICMS do Ativo Permanente – CIAP
'BLOCO G – CONTROLE DO CRÉDITO DE ICMS DO ATIVO PERMANENTE CIAP
'REGISTRO G001: ABERTURA DO BLOCO G
'REGISTRO G990: ENCERRAMENTO DO BLOCO G

Dim rsValG110_Ant As DAO.Recordset
Set rsValG110_Ant = Db.OpenRecordset("SELECT Sum(tbImobilizado.Valor_ICMS) AS Valor_ICMS FROM tbImobilizado " & _
"WHERE tbImobilizado.DataEmissao>=#" & cDtIniVb & "# and tbImobilizado.DataEmissao<=#" & cDtFimVb & "#  ;")

Dim rsValG110_Total As DAO.Recordset
Set rsValG110_Total = Db.OpenRecordset("SELECT Sum(tbImobilizado.Valor_ICMS_total) AS Valor_ICMS_Total FROM tbImobilizado " & _
"WHERE tbImobilizado.DataEmissao>=#" & cDtIniVb & "# and tbImobilizado.DataEmissao<=#" & cDtFimVb & "#  ;")

'Aqui entende-se saidas como vendas!!!
Dim rsValG110_Saidas As DAO.Recordset
Set rsValG110_Saidas = Db.OpenRecordset("SELECT Sum(tbVendasDet.ValorTot) AS ValorTot " & _
"FROM tbVendas INNER JOIN tbVendasDet ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao)>=#" & cDtIniVb & "# And (tbVendas.DataEmissao)<=#" & cDtFimVb & "#));")

Dim rsValG110_Inicial As DAO.Recordset
Set rsValG110_Inicial = Db.OpenRecordset("select SUM(Valor_ICMS_Total) AS Valor_ICMS_Totals from tbImobilizado " & _
"where ano = '" & year(cDtIniVb) & "' and mes = " & month(cDtFimVb) & "  and ciclo = '1';")


Dim rsG125 As DAO.Recordset
'Set rsG125 = db.OpenRecordset("SELECT tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.IDProd, First(tbImobilizado.DataEmissao) AS DataEmissao, Sum(tbImobilizado.Valor_ICMS) AS Valor_ICMS, Sum(tbImobilizado.Valor_ICMS_ST) AS Valor_ICMS_ST, First(tbImobilizado.Ciclo) AS Ciclo " & _
'"FROM tbImobilizado " & _
'"WHERE (((tbImobilizado.DataEmissao) >= #" & cDtINI & "# And (tbImobilizado.DataEmissao) <= # " & cDtFIM & "#)) " & _
'"GROUP BY tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.IDProd;")

Set rsG125 = Db.OpenRecordset("SELECT tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.IDProd, First(tbImobilizado.DataEmissao) AS DataEmissao, sum(tbImobilizado.Qnt) as Qnt, Sum(tbImobilizado.Valor_ICMS_total) AS Valor_ICMS_total, Sum(tbImobilizado.Valor_ICMS) AS Valor_ICMS, Sum(tbImobilizado.Valor_ICMS_ST) AS Valor_ICMS_ST, First(tbImobilizado.Ciclo) AS Ciclo, tbImobilizadoCadastro.Bem_Componente " & _
"FROM tbImobilizado left outer join tbImobilizadoCadastro " & _
"on tbImobilizado.IDProd = tbImobilizadoCadastro.IDProd " & _
"WHERE tbImobilizado.DataEmissao >= #" & cDtIniVb & "# And tbImobilizado.DataEmissao <= # " & cDtFimVb & "# " & _
"GROUP BY tbImobilizado.ANO, tbImobilizado.MES, tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente; ")

'AQUI TEM QUE TER O FIRST PORQUE QUANDO TEM IMOBILIZAÇÃO INICIAL DE UM ITEM QUE JÁ EXISTE DÁ ERRO NO PG

'"WHERE (((tbImobilizado.DataEmissao) >= #" & cDtINI & "# And (tbImobilizado.DataEmissao) <= # " & cDtFIM & "#)) " & _



Dim rsG130 As DAO.Recordset
Dim rsG140 As DAO.Recordset

Dim totalIcms As Double
totalIcms = 0
Do Until rsG125.EOF
If rsG125!Ciclo = 1 Then
totalIcms = totalIcms + rsG125!Valor_ICMS_total
Else
End If
rsG125.MoveNext
Loop

rsG125.MoveFirst


Dim cLinG As Integer
cLinG = 0
'cG001
cG001 = "|" & "G001" & "|" & "0" & "|"
cLinG = cLinG + 1
lG001 = lG001 + 1
Print #iArq, cG001


'cG110
If IsNull(rsValG110_Inicial!Valor_ICMS_Totals) Then
cG110 = "|" & "G110" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|" & "0" & "|" & Round(rsValG110_Ant!Valor_ICMS, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
Else
cG110 = "|" & "G110" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|" & Round(totalIcms, 2) & "|" & Round(rsValG110_Ant!Valor_ICMS, 2) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & "0" & "|"
End If
cLinG = cLinG + 1
lG110 = lG110 + 1
Print #iArq, cG110

'cG125

Do Until rsG125.EOF
Select Case rsG125!Ciclo
Case Is = "1"
cCiclo = "SI"
Case Is = "49"
cCiclo = "BA"
Case Else
    Select Case rsG125!Bem_Componente
    Case Is = "BEM"
    cCiclo = "IM"
    Case Is = "COMP"
    cCiclo = "IA"
    Case Else
    cCiclo = "IA"
    End Select
'cCiclo = "SI"
End Select

If rsG125!Valor_ICMS = 0 Then
GoTo prox_Lin_CIAP
Else
End If

'cG125 = "|" & "G125" & "|" & rsG125!IDProd & "|" & Replace(rsG125!DataEmissao, "/", "") & "|" & cCiclo & "|" & Round(rsG125!Valor_ICMS_total, 2) & "|" & Round(rsG125!Valor_ICMS_ST, 2) & "|" & "0" & "|" & "0" & "|" & rsG125!Ciclo & "|" & Round(rsG125!Valor_ICMS, 2) & "|"
cG125 = "|" & "G125" & "|" & rsG125!IDProd & "|" & cSTR_DtINI & "|" & cCiclo & "|" & Round(rsG125!Valor_ICMS_total, 2) & "|" & Round(rsG125!Valor_ICMS_ST, 2) & "|" & "0" & "|" & "0" & "|" & rsG125!Ciclo & "|" & Round(rsG125!Valor_ICMS, 2) & "|"
Print #iArq, cG125
cLinG = cLinG + 1
lG125 = lG125 + 1
'cG130 - subnivel do G125
'Set rsG130 = db.OpenRecordset("select * from tbcompras where chaveNF = '" & rsG125!chaveNFE & "'")

Set rsG130 = Db.OpenRecordset("SELECT tbComprasDet.IDProd, tbCompras.DataEmissao, tbCompras.IdFornecedor, tbCompras.Serie, tbCompras.NumNF, tbCompras.ChaveNF " & _
"FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE (((tbComprasDet.IDProd)=" & rsG125!IDProd & ")); ")

cG130 = "|" & "G130" & "|" & "1" & "|" & rsG130!IdFornecedor & "|" & "55" & "|" & rsG130!Serie & "|" & rsG130!NumNF & "|" & rsG130!chavenf & "|" & Replace(rsG130!DataEmissao, "/", "") & "|" & "|"
Print #iArq, cG130
cLinG = cLinG + 1
lG130 = lG130 + 1
'Set rsG140 = DB.OpenRecordset("select * from tbcomprasDet where IDCompra = '" & rsG130!ID & "' and IDProd = " & rsG125!IDProd & "")
Set rsG140 = Db.OpenRecordset("select q1.IDProd, q2.Unid from tbComprasDet q1 inner join tbCadProd q2 on q1.IDProd = q2.IDProd where q1.IDProd = " & rsG125!IDProd & "")
'precisa colocar ainda o ICMS do frete e o ICMS do Difal.
cG140 = "|" & "G140" & "|" & "1" & "|" & rsG125!IDProd & "|" & Replace(rsG125!Qnt, ".", ",") & "|" & rsG140!Unid & "|" & Round(rsG125!Valor_ICMS_total, 2) & "|" & Round(rsG125!Valor_ICMS_ST, 2) & "|" & "0" & "|" & "0" & "|"
Print #iArq, cG140
cLinG = cLinG + 1
lG140 = lG140 + 1
rsG130.Close
'rsG140.Close
prox_Lin_CIAP:
rsG125.MoveNext

Loop

'cG990
cLinG = cLinG + 1
lG990 = lG990 + 1
cG990 = "|" & "G990" & "|" & cLinG & "|"
Print #iArq, cG990



'H - Inventário Físico
'BLOCO H: INVENTÁRIO FÍSICO
'REGISTRO H005: TOTAIS DO INVENTÁRIO
'REGISTRO H010: INVENTÁRIO.
'REGISTRO H990: ENCERRAMENTO DO BLOCO H.



Dim rsIventario As DAO.Recordset
Set rsIventario = Db.OpenRecordset("select * from tbIventario where id = " & cIDIventario & " ")

Dim rsIveDet As DAO.Recordset
'Set rsIveDet = db.OpenRecordset("select * from tbIventarioDet where id_iventario = " & cIDIventario & " ")
Set rsIveDet = Db.OpenRecordset("SELECT tbIventarioDet.ID_Iventario, tbIventarioDet.ID_Prod AS ID_Prod, tbIventarioDet.Unid_Item AS Unid_Item, Sum(tbIventarioDet.Qtd) AS Qtd, Avg(tbIventarioDet.Valor_Unit) AS Valor_Unit, Sum(tbIventarioDet.Valor_Total) AS Valor_Total, Sum(tbIventarioDet.Valor_Item_IR) AS Valor_Item_IR " & _
"FROM tbIventarioDet " & _
"GROUP BY tbIventarioDet.ID_Iventario, tbIventarioDet.ID_Prod, tbIventarioDet.Unid_Item " & _
"HAVING (((tbIventarioDet.ID_Iventario)=" & cIDIventario & "));")


Dim IvenData As String
Dim cLinH As Integer

If rsIveDet.RecordCount = 0 Then
'sem iventario
IvenData = "NAO"
    
cLinH = 0
'cH001
cH001 = "|" & "H001" & "|" & "1" & "|"
cLinH = cLinH + 1
lH001 = lH001 + 1
Print #iArq, cH001


Else

'com iventario
IvenData = "SIM"

cLinH = 0
'cH001
cH001 = "|" & "H001" & "|" & "0" & "|"
cLinH = cLinH + 1
lH001 = lH001 + 1
Print #iArq, cH001


'cH005
Do Until rsIventario.EOF
cH005 = "|" & "H005" & "|" & Replace(rsIventario!DataIventario, "/", "") & "|" & Round(rsIventario!ValorTotalEstoque, 2) & "|" & "01" & "|"
cLinH = cLinH + 1
lH005 = lH005 + 1
Print #iArq, cH005
rsIventario.MoveNext
Loop

'cH010

Do Until rsIveDet.EOF
cH010 = "|" & "H010" & "|" & rsIveDet!ID_Prod & "|" & rsIveDet!Unid_Item & "|" & rsIveDet!Qtd & "|" & Round(rsIveDet!Valor_Unit, 2) & "|" & Round(rsIveDet!Valor_Total, 2) & "|" & "0" & "|" & "|" & "|" & "1" & "|" & Round(rsIveDet!Valor_Item_IR, 2) & "|"
Print #iArq, cH010
cLinH = cLinH + 1
lH010 = lH010 + 1
rsIveDet.MoveNext
Loop

End If

'cH990
cLinH = cLinH + 1
lH990 = lH990 + 1
cH990 = "|" & "H990" & "|" & cLinH & "|"
Print #iArq, cH990



'K** -Controle da Produção e do Estoque
Dim cLinK As Integer
cLinK = 0

Dim rsConsumo As DAO.Recordset
'Set rsConsumo = db.OpenRecordset("SELECT tb_Registro_Consumo.ID_Produto, tb_Registro_Consumo.Desc_Produto, tbCadProd.CONSUMO, tbCadProd.EMBALAGEM, tbCadProd.MAT_PRIMA, tb_Registro_Consumo.QT_CONSUMO, tb_Registro_Producao.Data_Brassagem " & _
'"FROM tb_Registro_Producao INNER JOIN (tbCadProd INNER JOIN tb_Registro_Consumo ON tbCadProd.IDProd = tb_Registro_Consumo.ID_Produto) ON tb_Registro_Producao.ID = tb_Registro_Consumo.ID_Lote " & _
'"WHERE (((tb_Registro_Producao.Data_Brassagem)>=#" & cDtINI & "# And (tb_Registro_Producao.Data_Brassagem)<=#" & cDtFIM & "#));")
''"(((tbCadProd.IDProd)<>2800) AND ((tbCadProd.CONSUMO)='SIM') AND ((tbCompras.DataEmissao)>=#" & cDtINI & "# And (tbCompras.DataEmissao)<=#" & cDtFIM & "#)) OR " & _

Dim rsProducao As DAO.Recordset
Set rsProducao = Db.OpenRecordset("SELECT tb_Registro_Producao.ID, tb_Registro_Envase.ID_Produto, tb_Registro_Envase.Data, tb_Registro_Envase.Qt " & _
"FROM tb_Registro_Producao INNER JOIN tb_Registro_Envase ON tb_Registro_Producao.ID = tb_Registro_Envase.ID " & _
"WHERE (((tb_Registro_Envase.Data)>=#" & cDtIni & "# And (tb_Registro_Envase.Data)<=#" & cDtFim & "#));")





'VENDAS PRODUCAO
Dim rsSTK_1 As DAO.Recordset
'Set rsSTK_1 = db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.PROD_FINAL, tbVendas.DataEmissao, Sum(tbVendasDet.Qnt) AS QT_Vendida " & _
"FROM tbVendas RIGHT JOIN (tbCadProd LEFT JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.PROD_FINAL, tbVendas.DataEmissao " & _
"HAVING (((tbCadProd.PROD_FINAL)='SIM') AND ((tbVendas.DataEmissao)>=#" & cDtINI & "# And (tbVendas.DataEmissao)<=#" & cDtFIM & "#));")

Set rsSTK_1 = Db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.PROD_FINAL, Sum(tbVendasDet.Qnt) AS QT_Vendida " & _
"FROM tbVendas RIGHT JOIN (tbCadProd LEFT JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE (((tbVendas.DataEmissao) >= #" & cDtIni & "# And (tbVendas.DataEmissao) <= #" & cDtFim & "#)) " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.PROD_FINAL " & _
"HAVING (((tbCadProd.PROD_FINAL)='SIM'));")


'VENDAS DE REVENDA
Dim rsSTK_2 As DAO.Recordset
Set rsSTK_2 = Db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.REVENDA, Sum(tbVendasDet.Qnt) AS QT_Vendida " & _
"FROM tbVendas RIGHT JOIN (tbCadProd LEFT JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda " & _
"WHERE ((tbVendas.DataEmissao)>=#" & cDtIni & "# And (tbVendas.DataEmissao)<=#" & cDtFim & "#) " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.REVENDA " & _
"HAVING (((tbCadProd.REVENDA)='SIM'));")

'COMPRAS DE CONSUMO
'CONSUMO NÃO ENTRA NO REGISTRO K100 - APURACAO DO ICMS E IPI PORQUE NÃO DÁ DIREITO A CREDITO
'Dim rsSTK_3 As DAO.Recordset
'Set rsSTK_3 = db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.CONSUMO, tbCompras.DataEmissao, Sum(tbComprasDet.Qnt) AS QT_Comprada " & _
'"FROM tbCompras INNER JOIN ((tbVendas RIGHT JOIN (tbCadProd LEFT JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) " & _
'"INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
'"GROUP BY tbCadProd.IDProd, tbCadProd.CONSUMO, tbCompras.DataEmissao " & _
'"HAVING (((tbCadProd.CONSUMO)='SIM') AND ((tbCompras.DataEmissao)>=#" & cDtINI & "# And (tbCompras.DataEmissao)<=#" & cDtFIM & "#)) ;")


'COMPRAS DE EMBALAGEM
Dim rsSTK_4 As DAO.Recordset
Set rsSTK_4 = Db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.EMBALAGEM, Sum(tbComprasDet.Qnt) AS QT_Comprada " & _
"FROM tbCompras INNER JOIN ((tbVendas RIGHT JOIN (tbCadProd LEFT JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) " & _
"INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE ((tbCompras.DataEmissao)>=#" & cDtIni & "# And (tbCompras.DataEmissao)<=#" & cDtFim & "#) " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.EMBALAGEM " & _
"HAVING (((tbCadProd.EMBALAGEM)='SIM'));")

'COMPRAS DE MATERIA PRIMA
Dim rsSTK_5 As DAO.Recordset
Set rsSTK_5 = Db.OpenRecordset("SELECT tbCadProd.IDProd, tbCadProd.MAT_PRIMA, Sum(tbComprasDet.Qnt) AS QT_Comprada " & _
"FROM tbCompras INNER JOIN ((tbVendas RIGHT JOIN (tbCadProd LEFT JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda) " & _
"INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra " & _
"WHERE ((tbCompras.DataEmissao)>=#" & cDtIni & "# And (tbCompras.DataEmissao)<=#" & cDtFim & "#) " & _
"GROUP BY tbCadProd.IDProd, tbCadProd.MAT_PRIMA " & _
"HAVING (((tbCadProd.MAT_PRIMA)='SIM'));")




'REGISTRO K001: ABERTURA DO BLOCO K
'cK001
cK001 = "|" & "K001" & "|" & "0" & "|"
cLinK = cLinK + 1
lK001 = lK001 + 1
Print #iArq, cK001

'REGISTRO K010: INFORMA��O SOBRE O TIPO DE LEIAUTE (SIMPLIFICADO / COMPLETO)
cK010 = "|" & "K010" & "|" & "1" & "|"
cLinK = cLinK + 1
lK010 = lK010 + 1
Print #iArq, cK010

'REGISTRO K100: PERÍODO DE APURAÇÃO DO ICMS/IPI
'cK100
cK100 = "|" & "K100" & "|" & cSTR_DtINI & "|" & cSTR_DtFIM & "|"
cLinK = cLinK + 1
lK100 = lK100 + 1
Print #iArq, cK100

Dim cDtK200 As String
cDtK200 = "NAO"

'REGISTRO K200: ESTOQUE ESCRITURADO
'cK200
'STK_1
Do Until rsSTK_1.EOF
cK200 = "|" & "K200" & "|" & cSTR_DtFIM & "|" & rsSTK_1!IDProd & "|" & Round(rsSTK_1!Qt_Vendida * 0.8, 0) & "|" & "0" & "|" & "|"
rsSTK_1.MoveNext
cLinK = cLinK + 1
lK200 = lK200 + 1
Print #iArq, cK200
cDtK200 = "SIM"
Loop
'STK_2
Do Until rsSTK_2.EOF
cK200 = "|" & "K200" & "|" & cSTR_DtFIM & "|" & rsSTK_2!IDProd & "|" & Round(rsSTK_2!Qt_Vendida * 0.9, 0) & "|" & "0" & "|" & "|"
rsSTK_2.MoveNext
cLinK = cLinK + 1
lK200 = lK200 + 1
Print #iArq, cK200
cDtK200 = "SIM"
Loop
'STK_3
'ITENS DE CONSUMO NÃO DÁ DIREITO A CREDITO DE ICMS E IPI
'Do Until rsSTK_3.EOF
'If rsSTK_3!IDProd = 2800 Then
'rsSTK_3.MoveNext
'Else
'cK200 = "|" & "K200" & "|" & cSTR_DtFIM & "|" & rsSTK_3!IDProd & "|" & Round(rsSTK_3!Qt_Comprada * 0.15, 0) & "|" & "0" & "|" & "|"
'rsSTK_3.MoveNext
'cLinK = cLinK + 1
'lK200 = lK200 + 1
'Print #iArq, cK200
'cDtK200 = "SIM"
'End If
'Loop
'STK_4
Do Until rsSTK_4.EOF
cK200 = "|" & "K200" & "|" & cSTR_DtFIM & "|" & rsSTK_4!IDProd & "|" & Round(rsSTK_4!Qt_Comprada * 0.3, 0) & "|" & "0" & "|" & "|"
rsSTK_4.MoveNext
cLinK = cLinK + 1
lK200 = lK200 + 1
Print #iArq, cK200
cDtK200 = "SIM"
Loop
'STK_5
Do Until rsSTK_5.EOF
If rsSTK_5!IDProd = 2800 Then
rsSTK_5.MoveNext
Else
cK200 = "|" & "K200" & "|" & cSTR_DtFIM & "|" & rsSTK_5!IDProd & "|" & Round(rsSTK_5!Qt_Comprada * 0.1, 0) & "|" & "0" & "|" & "|"
rsSTK_5.MoveNext
cLinK = cLinK + 1
lK200 = lK200 + 1
Print #iArq, cK200
cDtK200 = "SIM"
End If
Loop


'REGISTRO K230: ITENS PRODUZIDOS
'cK230
'PRIMEIRO FAZ O RATEIO DO CONSUMO NA tb_Registro_Envase_Rateio_Temp
strSQL = ("DELETE FROM tb_Registro_Envase_Rateio_Temp")

Call ConnectToDataBase
Conn.Execute strSQL
strSQL = ("INSERT INTO tb_Registro_Envase_Rateio_Temp ( Data, ID, ID_Produto, Unid_Med, Qt, VOL_LIT, PART ) " & _
"SELECT tb_Registro_Envase.Data, tb_Registro_Envase.ID, tb_Registro_Envase.ID_Produto, tb_Registro_Envase.Unid_Med, tb_Registro_Envase.Qt, LITROS*QT AS VOL_LIT, 0 AS PART " & _
"FROM tbCadProd INNER JOIN tb_Registro_Envase ON tbCadProd.IDProd = tb_Registro_Envase.ID_Produto " & _
"WHERE (((tb_Registro_Envase.Data)>='" & cDtIni & "' And (tb_Registro_Envase.Data)<='" & cDtFim & "') AND ((tbCadProd.PROD_FINAL)='SIM'));")
Conn.Execute strSQL
'CALCULA O PESO, RATEIO DE CADA PRODUTO NO CONSUMO
Dim rsRateio As DAO.Recordset
Set rsRateio = Db.OpenRecordset("SELECT tb_Registro_Envase_Rateio_Temp.ID, Sum(tb_Registro_Envase_Rateio_Temp.VOL_LIT) AS VOL_LIT_TOT " & _
"FROM tb_Registro_Envase_Rateio_Temp " & _
"GROUP BY tb_Registro_Envase_Rateio_Temp.ID;")
Dim rsRateioConsumo As DAO.Recordset


Do Until rsRateio.EOF
    Set rsRateioConsumo = Db.OpenRecordset("select * from tb_Registro_Envase_Rateio_Temp where id = " & rsRateio!ID & "")
    Do Until rsRateioConsumo.EOF
    rsRateioConsumo.Edit
    rsRateioConsumo!Part = rsRateioConsumo!VOL_LIT / rsRateio!VOL_LIT_TOT
    rsRateioConsumo.Update
    rsRateioConsumo.MoveNext
    Loop
    rsRateio.MoveNext
Loop

'REGISTRO K230: ITENS PRODUZIDOS
'cK230

Dim cTemproducao As String

If rsProducao.RecordCount > 0 Then
rsProducao.MoveFirst
cTemproducao = "SIM"
Do Until rsProducao.EOF
cK230 = "|" & "K230" & "|" & Replace(Format(rsProducao!Data, "dd/mm/yyyy"), "/", "") & "|" & Replace(Format(rsProducao!Data, "dd/mm/yyyy"), "/", "") & "|" & rsProducao!ID & "|" & rsProducao!ID_Produto & "|" & rsProducao!Qt & "|"
cLinK = cLinK + 1
lK230 = lK230 + 1
Print #iArq, cK230


'REGISTRO K235: INSUMOS CONSUMIDOS
'cK235
Set rsConsumo = Db.OpenRecordset("SELECT tb_Registro_Envase_Rateio_Temp.ID, tb_Registro_Envase_Rateio_Temp.ID_Produto, tb_Registro_Consumo.ID_Produto AS ID_PROD_CONSUMO, Round(PART*QT_CONSUMO,2) AS CONSUMO, tb_Registro_Envase_Rateio_Temp.Data " & _
"FROM tb_Registro_Envase_Rateio_Temp INNER JOIN tb_Registro_Consumo ON tb_Registro_Envase_Rateio_Temp.ID = tb_Registro_Consumo.ID_Lote " & _
"WHERE (((tb_Registro_Envase_Rateio_Temp.ID)=" & rsProducao!ID & ") AND ((tb_Registro_Envase_Rateio_Temp.ID_Produto)=" & rsProducao!ID_Produto & ") AND ((Round(PART*QT_CONSUMO,2))>0) AND ((tb_Registro_Envase_Rateio_Temp.Data)>=#" & cDtIni & "# And (tb_Registro_Envase_Rateio_Temp.Data)<=#" & cDtFim & "#));")


rsConsumo.MoveFirst
Do Until rsConsumo.EOF
cK235 = "|" & "K235" & "|" & Replace(Format(rsProducao!Data, "dd/mm/yyyy"), "/", "") & "|" & rsConsumo!ID_PROD_CONSUMO & "|" & rsConsumo!CONSUMO & "|" & "|"
cLinK = cLinK + 1
lK235 = lK235 + 1
Print #iArq, cK235
rsConsumo.MoveNext
Loop

rsProducao.MoveNext
Loop
    
Else
cTemproducao = "NAO"
End If


'REGISTRO K990: ENCERRAMENTO DO BLOCO K
'cK990
cLinK = cLinK + 1
lK990 = lK990 + 1
cK990 = "|" & "K990" & "|" & cLinK & "|"
Print #iArq, cK990

'1 - Outras Informações
Dim cLin1 As Integer
cLin1 = 0

'REGISTRO 1001: ABERTURA DO BLOCO 1
'c1001
c1001 = "|" & "1001" & "|" & "0" & "|"
Print #iArq, c1001
cLin1 = cLin1 + 1
l1001 = l1001 + 1

'REGISTRO 1010: OBRIGATORIEDADE DE REGISTROS DO BLOCO 1
'c1010
c1010 = "|" & "1010" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|" & "N" & "|"
Print #iArq, c1010
cLin1 = cLin1 + 1
l1010 = l1010 + 1

'REGISTRO 1100: REGISTRO DE INFORMAÇÕES SOBRE EXPORTAÇÃO.
'Omitido

'REGISTRO 1105: DOCUMENTOS FISCAIS DE EXPORTAÇÃO.
'Omitido

'REGISTRO 1200: CONTROLE DE CRÉDITOS FISCAIS - ICMS.
'c1200

'ta estranho esse bloco porque ele pede credito de ICMS ST no código
'SP099719   ???


'Select Case rsResumoICMS!SALDO
'Case Is < 0
'c1200 = "|" & "1200" & "|" & "SP099719" & "|" & RoundDown(rsResumoICMS!CRED_MES_ANT, 2) & "|" & RoundDown(rsResumoICMS!CRED, 2) & "|" & "0" & "|" & RoundDown(rsResumoICMS!DEB + rsResumoICMS!SALDO, 2) & "|" & 0 & "|"
'Case Is >= 0
'c1200 = "|" & "1200" & "|" & "SP099719" & "|" & RoundDown(rsResumoICMS!CRED_MES_ANT, 2) & "|" & RoundDown(rsResumoICMS!CRED, 2) & "|" & "0" & "|" & RoundDown(rsResumoICMS!DEB, 2) & "|" & RoundDown(rsResumoICMS!CRED_TRANSPORTAR, 2) & "|"
'End Select
'cLin1 = cLin1 + 1
'l1200 = l1200 + 1
'Print #iArq, c1200

'REGISTRO 1210: UTILIZAÇÃO DE CRÉDITOS FISCAIS – ICMS.
'Select Case rsResumoICMS!DEB
'Case Is >= 0
'c1210 = "|" & "1210" & "|" & "SP01" & "|" & "NF-E VENDA" & "|" & Round(rsResumoICMS!DEB + rsResumoICMS!SALDO, 2) & "|" & "|"
'Case Is < 0
'c1210 = "|" & "1210" & "|" & "SP01" & "|" & "NF-E VENDA" & "|" & Round(rsResumoICMS!DEB + rsResumoICMS!SALDO, 2) & "|" & "|"
'End Select
'cLin1 = cLin1 + 1
'l1210 = l1210 + 1
'Print #iArq, c1210

'REGISTRO 1600: TOTAL DAS OPERAÇÕES COM CARTÃO DE CRÉDITO E/OU DÉBITO

'REGISTRO 1990: ENCERRAMENTO DO BLOCO 1
'c1990
cLin1 = cLin1 + 1
l1990 = l1990 + 1
c1990 = "|" & "1990" & "|" & cLin1 & "|"
Print #iArq, c1990

'9 - Controle e Encerramento do Arquivo Digital
'BLOCO 9: CONTROLE E ENCERRAMENTO DO ARQUIVO DIGITAL
'REGISTRO 9001: ABERTURA DO BLOCO 9
Dim cLin9 As Integer
cLin9 = 0

'c9001
c9001 = "|" & "9001" & "|" & "0" & "|"
Print #iArq, c9001
cLin9 = cLin9 + 1
l9001 = l9001 + 1
l99 = l99 + 1

'REGISTRO 9900: REGISTROS DO ARQUIVO.
'c9900
cLin9 = cLin9 + 1
l9900 = l9900 + 1
 'BLOCO 0
'0000
c9900 = "|" & "9900" & "|" & "0000" & "|" & l0000 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0001
c9900 = "|" & "9900" & "|" & "0001" & "|" & l0001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0002
c9900 = "|" & "9900" & "|" & "0002" & "|" & l0002 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0005
c9900 = "|" & "9900" & "|" & "0005" & "|" & l0005 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0100
c9900 = "|" & "9900" & "|" & "0100" & "|" & l0100 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0150
c9900 = "|" & "9900" & "|" & "0150" & "|" & l0150 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0190
c9900 = "|" & "9900" & "|" & "0190" & "|" & l0190 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0200
c9900 = "|" & "9900" & "|" & "0200" & "|" & l0200 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0300
c9900 = "|" & "9900" & "|" & "0300" & "|" & l0300 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0305
c9900 = "|" & "9900" & "|" & "0305" & "|" & l0305 & "|"
Print #iArq, c9900
l99 = l99 + 1
    '0400
    If c0400Data = "SIM" Then
    c9900 = "|" & "9900" & "|" & "0400" & "|" & l0400 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
'0500
c9900 = "|" & "9900" & "|" & "0500" & "|" & l0500 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0600
c9900 = "|" & "9900" & "|" & "0600" & "|" & l0600 & "|"
Print #iArq, c9900
l99 = l99 + 1
'0990
c9900 = "|" & "9900" & "|" & "0990" & "|" & l0990 & "|"
Print #iArq, c9900
l99 = l99 + 1

'BLOCO B
'B001
c9900 = "|" & "9900" & "|" & "B001" & "|" & lB001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'B990
c9900 = "|" & "9900" & "|" & "B990" & "|" & lB990 & "|"
Print #iArq, c9900
l99 = l99 + 1
'BLOCO B

''BLOCO C
'C001
c9900 = "|" & "9900" & "|" & "C001" & "|" & lC001 & "|"
Print #iArq, c9900
l99 = l99 + 1
    'C100
    If cDtC100 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "C100" & "|" & lC100 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
    'C170
    If cDtC100 = "SIM" And clin170 > 0 Then
    c9900 = "|" & "9900" & "|" & "C170" & "|" & lC170 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
'C190
    If cDtC100 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "C190" & "|" & lC190 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
'C500
    If cDtC500 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "C500" & "|" & lC500 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
'C510
'c9900 = "|" & "9900" & "|" & "C510" & "|" & lC510 & "|"
'Print #iArq, c9900
'l99 = l99 + 1

'C590
    If cDtC500 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "C590" & "|" & lC590 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
    
'C800
    If cDtC500 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "C800" & "|" & lC800 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If

'C810
    If cDtC500 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "C810" & "|" & lC810 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
'C850
    If cDtC500 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "C850" & "|" & lC850 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If

'C990
c9900 = "|" & "9900" & "|" & "C990" & "|" & lC990 & "|"
Print #iArq, c9900
l99 = l99 + 1


''BLOCO D
'D001
c9900 = "|" & "9900" & "|" & "D001" & "|" & lD001 & "|"
Print #iArq, c9900
l99 = l99 + 1
    If flagTrp = 1 Then
        Else
        'D100
        c9900 = "|" & "9900" & "|" & "D100" & "|" & lD100 & "|"
        Print #iArq, c9900
        l99 = l99 + 1
        'D190
        c9900 = "|" & "9900" & "|" & "D190" & "|" & lD190 & "|"
        Print #iArq, c9900
        l99 = l99 + 1
    End If

'D990
c9900 = "|" & "9900" & "|" & "D990" & "|" & lD990 & "|"
Print #iArq, c9900
l99 = l99 + 1
''BLOCO E
'E001
c9900 = "|" & "9900" & "|" & "E001" & "|" & lE001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'E100
c9900 = "|" & "9900" & "|" & "E100" & "|" & lE100 & "|"
Print #iArq, c9900
l99 = l99 + 1
'E110
c9900 = "|" & "9900" & "|" & "E110" & "|" & lE110 & "|"
Print #iArq, c9900
l99 = l99 + 1
'E116
c9900 = "|" & "9900" & "|" & "E116" & "|" & lE116 & "|"
Print #iArq, c9900
l99 = l99 + 1
'E500
c9900 = "|" & "9900" & "|" & "E500" & "|" & lE500 & "|"
Print #iArq, c9900
l99 = l99 + 1
'E510
c9900 = "|" & "9900" & "|" & "E510" & "|" & lE510 & "|"
Print #iArq, c9900
l99 = l99 + 1
'E520
c9900 = "|" & "9900" & "|" & "E520" & "|" & lE520 & "|"
Print #iArq, c9900
l99 = l99 + 1
'E990
c9900 = "|" & "9900" & "|" & "E990" & "|" & lE990 & "|"
Print #iArq, c9900
l99 = l99 + 1
''BLOCO G
'G001
c9900 = "|" & "9900" & "|" & "G001" & "|" & lG001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'G110
c9900 = "|" & "9900" & "|" & "G110" & "|" & lG110 & "|"
Print #iArq, c9900
l99 = l99 + 1
'G125
c9900 = "|" & "9900" & "|" & "G125" & "|" & lG125 & "|"
Print #iArq, c9900
l99 = l99 + 1
'G130
c9900 = "|" & "9900" & "|" & "G130" & "|" & lG130 & "|"
Print #iArq, c9900
l99 = l99 + 1
'G140
c9900 = "|" & "9900" & "|" & "G140" & "|" & lG140 & "|"
Print #iArq, c9900
l99 = l99 + 1
'G990
c9900 = "|" & "9900" & "|" & "G990" & "|" & lG990 & "|"
Print #iArq, c9900
l99 = l99 + 1
''BLOCO H
'H001
c9900 = "|" & "9900" & "|" & "H001" & "|" & lH001 & "|"
Print #iArq, c9900
l99 = l99 + 1

If IvenData = "SIM" Then
'H005
c9900 = "|" & "9900" & "|" & "H005" & "|" & lH005 & "|"
Print #iArq, c9900
l99 = l99 + 1
'H010
c9900 = "|" & "9900" & "|" & "H010" & "|" & lH010 & "|"
Print #iArq, c9900
l99 = l99 + 1
Else: End If

'H990
c9900 = "|" & "9900" & "|" & "H990" & "|" & lH990 & "|"
Print #iArq, c9900
l99 = l99 + 1
''BLOCO K
'K001
c9900 = "|" & "9900" & "|" & "K001" & "|" & lK001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'K010
c9900 = "|" & "9900" & "|" & "K010" & "|" & lK010 & "|"
Print #iArq, c9900
l99 = l99 + 1

'K100
c9900 = "|" & "9900" & "|" & "K100" & "|" & lK100 & "|"
Print #iArq, c9900
l99 = l99 + 1
'K200
    If cDtK200 = "SIM" Then
    c9900 = "|" & "9900" & "|" & "K200" & "|" & lK200 & "|"
    Print #iArq, c9900
    l99 = l99 + 1
    Else: End If
    

'K230
If cTemproducao = "SIM" Then
c9900 = "|" & "9900" & "|" & "K230" & "|" & lK230 & "|"
Print #iArq, c9900
l99 = l99 + 1
'K235
c9900 = "|" & "9900" & "|" & "K235" & "|" & lK235 & "|"
Print #iArq, c9900
l99 = l99 + 1
    Else
    End If
    
'K990
c9900 = "|" & "9900" & "|" & "K990" & "|" & lK990 & "|"
Print #iArq, c9900
l99 = l99 + 1
''BLOCO 1
'1001
c9900 = "|" & "9900" & "|" & "1001" & "|" & l1001 & "|"
Print #iArq, c9900
l99 = l99 + 1
'1010
c9900 = "|" & "9900" & "|" & "1010" & "|" & l1010 & "|"
Print #iArq, c9900
l99 = l99 + 1
'1200
'c9900 = "|" & "9900" & "|" & "1200" & "|" & l1200 & "|"
'Print #iArq, c9900
'l99 = l99 + 1
    '1210
'    If l1210 > 0 Then
'    c9900 = "|" & "9900" & "|" & "1210" & "|" & l1210 & "|"
'    Print #iArq, c9900
'    l99 = l99 + 1
'    Else: End If
'1990
c9900 = "|" & "9900" & "|" & "1990" & "|" & l1990 & "|"
Print #iArq, c9900
l99 = l99 + 1
''BLOCO 9
'9001
c9900 = "|" & "9900" & "|" & "9001" & "|" & l9001 & "|"
Print #iArq, c9900
l99 = l99 + 1


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


'9990
'c9001 = "|" & "9900" & "|" & "9990" & "|" & "0" & "|"
'Print #iArq, c9001
'9999
'c9001 = "|" & "9999" & "|" & "9999" & "|" & "1" & "|"
'Print #iArq, c9001
l99 = l99 + 1
l99 = l99 + 1
'REGISTRO 9990: ENCERRAMENTO DO BLOCO 9
l99 = l99 + 1
c9990 = "|" & "9990" & "|" & l99 & "|"
Print #iArq, c9990


'REGISTRO 9999: ENCERRAMENTO DO ARQUIVO DIGITAL.

cTotal9999 = l0000 + l0001 + l0002 + l0005 + l0100 + l0150 + l0190 + l0200 + l0300 + l0305 + l0400 + l0500 + l0600 + l0990 + lB001 + lB990 + lC001 + lC100 + lC170 + lC190 + lC500 + lC590 + lC800 + lC810 + lC850 + lC990 + lD001 + lD100 + lD190 + lD990 + lE001 + lE100 + lE110 + lE116 + lE500 + lE510 + lE520 + lE990 + lG001 + lG110 + lG125 + lG130 + lG140 + lG990 + lH001 + lH005 + lH010 + lH990 + lK001 + lK010 + lK100 + lK200 + lK230 + lK235 + lK990 + l1001 + l1010 + l1200 + l1210 + l1990 + l9001 + l9900
c9999 = "|" & "9999" & "|" & cTotal9999 + l99 - 2 & "|"
Print #iArq, c9999

Call DisconnectFromDataBase
Close #iArq
'DoCmd.setwarnings (True)
MsgBox ("Arquivo Gerado")


End Function
