Attribute VB_Name = "modImportaXML_Venda"
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

Public Function ImportaXML_Venda(LocalXml As String)
'---------------------------------------------------------------'
'                Criado por Cezar Barreto                       '
'           Em 12/02/2020 para Overture                         '
'---------------------------------------------------------------'

Call ConnectToDataBase

Dim doc As DOMDocument
Dim xDet As IXMLDOMNodeList

Dim NomeArq As String
Dim Db As Database
Dim rsFor, rsProd, rsCompra, rsCompDet As DAO.Recordset
Dim xProd As String
Dim I, X, regAtual As Integer
DoCmd.OpenForm "frmAguarde"
Forms!frmaguarde!txt2 = "Xml Com Erros:"
Diret = LocalXml
NomeArq = Dir(Diret & "*.XML", vbArchive)

Dim cBaseCalcPis As String


Set Db = CurrentDb()
Set doc = New DOMDocument

contNF_Vend = 0
contProd = 0
contCli = 0



'Buscará todos os arquivos com extenção .xml da pasta selecionada
Do While NomeArq <> ""
doc.Load (LocalXml & NomeArq) 'Pega a Pasta e o Nome do primeiro arquivo....
'Verifica se o Arquivo foi aberto corretamente e se possui chave. Se possuir importa, se nao Pula pra o Proximo!
If doc.validate.errorCode = -1072897500 And (doc.getElementsByTagName("chNFe").length) Then
'Set xDet = doc.getElementsByTagName("det")
Set xDet = doc.getElementsByTagName("dest")
'------------------------------------------------------------------------'
'Insere os Dados do Fornecedor, se nao for Cadastrado.
Set rsFor = Db.OpenRecordset("tbcliente")




X = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & xDet.Item(0).childNodes(0).Text & "'"), 0)     'X buscará o fornecedor na tabela "tbFornecedores"
'x = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & doc.childNodes(2).getElementsByTagName("dest")(0) & "'"), 0) 'X buscará o fornecedor na tabela "tbFornecedores"


                                                    
                                                    
If doc.getElementsByTagName("CNPJ")(0).Text = Forms!frmXMLinput!txt_CNPJ_Empresa Then 'Se CNPJ do emissor for igual ao da empresa, irá cadastrar como venda
Else
GoTo proximoarquivo
End If



If X <= 0 Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o cliente
    rsFor.AddNew
        
        'On Error Resume Next
        rsFor!Tipo = "CLIENTE"
        'rsFor!CNPJ = doc.getElementsByTagName("CNPJ_dest")(0).Text
        rsFor!CNPJ = xDet.Item(0).childNodes(0).Text
        varCNPJ = xDet.Item(0).childNodes(0).Text
        
        rsFor!RazaoSocial = xDet.Item(0).childNodes(1).Text
        
       
        Select Case xDet.Item(0).childNodes(3).Text
        Case "1"
        rsFor!CRT = "CONTRIBUINTE ICMS"
        rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        Case "2"
        rsFor!CRT = "ISENTO ICMS"
        rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        Case "9"
        rsFor!CRT = "NAO CONTRIBUINTE"
        Case Else
        rsFor!CRT = "NAO INFORMADO"
        End Select
        
        rsFor!Email = xDet.Item(0).selectSingleNode("email").Text
        
        
        Set xDet = doc.getElementsByTagName("enderDest")
        
        rsFor!Logradouro = xDet.Item(0).selectSingleNode("xLgr").Text
        rsFor!Nro = xDet.Item(0).selectSingleNode("nro").Text
        rsFor!CEP = xDet.Item(0).selectSingleNode("CEP").Text
'        On Error Resume Next
        
        
        rsFor!compl = xDet.Item(0).selectSingleNode("xCpl").Text
        rsFor!Bairro = xDet.Item(0).selectSingleNode("xBairro").Text
        rsFor!UF = xDet.Item(0).selectSingleNode("UF").Text
        rsFor!Municipio = xDet.Item(0).selectSingleNode("xMun").Text
        rsFor!Pais = xDet.Item(0).selectSingleNode("xPais").Text
        'rsFor!fone = xDet.Item(0).selectSingleNode("fone").Text
    
        
        contCli = contCli + 1
        
        
    rsFor.Update
'Apos cadastrar o fornecedor, x buscara o ID desse fornecedor para ser utilizado na importação do xml em questao

'x = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & doc.getElementsByTagName("CNPJ_dest")(0).Text & "'"), 0)



X = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & varCNPJ & "'"), 0)     'X buscará o fornecedor na tabela "tbFornecedores"

Else
'atualiza dados do fornecedor já cadastrado
rsFor.Close
Set rsFor = Db.OpenRecordset("SELECT * FROM tbcliente WHERE idCliente = " & X & "")
rsFor.Edit
'On Error Resume Next
        
         rsFor!Tipo = "CLIENTE"
        'rsFor!CNPJ = doc.getElementsByTagName("CNPJ_dest")(0).Text
        rsFor!CNPJ = xDet.Item(0).childNodes(0).Text
        varCNPJ = xDet.Item(0).childNodes(0).Text
        
        rsFor!RazaoSocial = xDet.Item(0).childNodes(1).Text
        
       
        Select Case xDet.Item(0).childNodes(3).Text
        Case "1"
        rsFor!CRT = "CONTRIBUINTE ICMS"
        rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        Case "2"
        rsFor!CRT = "ISENTO ICMS"
        rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        Case "9"
        rsFor!CRT = "NAO CONTRIBUINTE"
        Case Else
        rsFor!CRT = "NAO INFORMADO"
        End Select
        
        rsFor!Email = xDet.Item(0).selectSingleNode("email").Text
        
        
        Set xDet = doc.getElementsByTagName("enderDest")
        
        rsFor!Logradouro = xDet.Item(0).selectSingleNode("xLgr").Text
        rsFor!Nro = xDet.Item(0).selectSingleNode("nro").Text
        rsFor!CEP = xDet.Item(0).selectSingleNode("CEP").Text
        
        
        On Error Resume Next
        rsFor!compl = xDet.Item(0).selectSingleNode("xCpl").Text
        rsFor!Bairro = xDet.Item(0).selectSingleNode("xBairro").Text
        rsFor!Fone = xDet.Item(0).selectSingleNode("fone").Text
        On Error GoTo 0
        
        rsFor!UF = xDet.Item(0).selectSingleNode("UF").Text
        rsFor!Municipio = xDet.Item(0).selectSingleNode("xMun").Text
        rsFor!Pais = xDet.Item(0).selectSingleNode("xPais").Text
 
On Error Resume Next
rsFor.Update

End If
rsFor.Close
Set rsFor = Nothing

'------------------------------------------------------------------------'
'Dados Principais da Nota de Venda (tbVendas)
Set xDet = doc.getElementsByTagName("det")
Set rsVenda = Db.OpenRecordset("tbVendas")

'verifica se o xml já foi processado antes pra não duplicar a linha
x1 = Nz(DLookup("ChaveNF", "tbVendas", "ChaveNF = '" & doc.getElementsByTagName("chNFe")(0).Text & "'"), 0)
If x1 = doc.getElementsByTagName("chNFe")(0).Text Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o fornecedor
GoTo proximoarquivo
Else
End If


rsVenda.AddNew
    rsVenda!IdCliente = X
    'Necessario essa verificação pois na versao XML 1.10 era somente Data (dEmi) ja na 3.0 mudou para DataHora (dhEmi)
    If (doc.getElementsByTagName("dhEmi").length) Then
    rsVenda!DataEmissao = Format(Left(doc.getElementsByTagName("dhEmi")(0).Text, 10), "dd/mm/yyyy")
    Else
    rsVenda!DataEmissao = Format(doc.getElementsByTagName("dEmi")(0).Text, "dd/mm/yyyy")
    End If
        'Passa os totais da NFe para a variavel xProd
        xProd = doc.getElementsByTagName("total")(0).XML
        'Valor Bruto sem desconto
       
       
    rsVenda!NumNF = doc.getElementsByTagName("nNF")(0).Text
    rsVenda!Serie = doc.getElementsByTagName("serie")(0).Text
    rsVenda!chavenf = doc.getElementsByTagName("chNFe")(0).Text
      
    rsVenda!Status = "ATIVO"
    
    Select Case doc.getElementsByTagName("tpNF")(0).Text
    Case 0
    rsVenda!TipoNF = "0-ENTRADA"
    Case 1
    rsVenda!TipoNF = "1-SAIDA"
    End Select
    
    rsVenda!NatOperacao = doc.getElementsByTagName("natOp")(0).Text
 '   rsVenda!ConsumidorFinal = doc.getElementsByTagName("indFinal")(0).Text
  '  rsVenda!DestOperacao = doc.getElementsByTagName("idDest")(0).Text
    
    Set xtotal = doc.getElementsByTagName("ICMSTot")
    
    rsVenda!VlrTotalProdutos = Replace(xtotal.Item(0).selectSingleNode("vProd").Text, ".", ",")
    rsVenda!VlrTotalFrete = Replace(xtotal.Item(0).selectSingleNode("vFrete").Text, ".", ",")
    rsVenda!VlrTotalSeguro = Replace(xtotal.Item(0).selectSingleNode("vSeg").Text, ".", ",")
    rsVenda!VlrDesconto = Replace(xtotal.Item(0).selectSingleNode("vDesc").Text, ".", ",")
    rsVenda!VlrDespesas = Replace(xtotal.Item(0).selectSingleNode("vOutro").Text, ".", ",")
    rsVenda!ICMS_BaseCalc = Replace(xtotal.Item(0).selectSingleNode("vBC").Text, ".", ",")
    rsVenda!ICMS_Valor = Replace(xtotal.Item(0).selectSingleNode("vICMS").Text, ".", ",")
    rsVenda!ICMS_ST_BaseCalc = Replace(xtotal.Item(0).selectSingleNode("vBCST").Text, ".", ",")
    rsVenda!ICMS_ST_Valor = Replace(xtotal.Item(0).selectSingleNode("vST").Text, ".", ",")
    rsVenda!IPI_Valor = Replace(xtotal.Item(0).selectSingleNode("vIPI").Text, ".", ",")
    rsVenda!PIS_Valor = Replace(xtotal.Item(0).selectSingleNode("vPIS").Text, ".", ",")
    rsVenda!COFINS_Valor = Replace(xtotal.Item(0).selectSingleNode("vCOFINS").Text, ".", ",")
    rsVenda!VlrTOTALNF = Replace(xtotal.Item(0).selectSingleNode("vNF").Text, ".", ",")
       
    contNF_Vend = contNF_Vend + 1
       
rsVenda.Update
regAtual = Nz(DLookup("ID", "tbVendas", "ChaveNF = '" & doc.getElementsByTagName("chNFe")(0).Text & "'"), 0)



rsVenda.Close
Set rsVenda = Nothing

'Registra no contas a receber
strSQL = ("INSERT INTO tb_Detalhe_Boletos_Vendas ( DtEmissao, ID_Cliente, Cliente, ValorOriginal, chave_NFe, STATUS, Duvidoso, NumBoleto ) " & _
                "SELECT tbVendas.DataEmissao, tbCliente.IDCliente, tbCliente.RazaoSocial, tbVendas.VlrTOTALNF, tbVendas.ChaveNF, 'ABERTO' AS STATUS, 'N' AS Duvidoso, 'LANC AUTOMATICO' AS boleto " & _
                "FROM tbCliente INNER JOIN tbVendas on tbCliente.IDCliente = tbVendas.IdCliente " & _
                "WHERE tbVendas.TipoNF='1-SAIDA' and chaveNF = '" & doc.getElementsByTagName("chNFe")(0).Text & "';")
Conn.Execute strSQL
'Registra no contas a receber

'------------------------------------------------------------------------'
' Dados dos Produtos

'verifica se o xml já foi processado antes pra não duplicar a linha
X = Nz(DLookup("IDVenda", "tbVendasDet", "IDVenda = " & regAtual & ""), 0)
If X = regAtual Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o fornecedor
GoTo proximoarquivo
Else
End If


I = 0
xProd = ""
'Aqui é o Loop que percorrerá pela Tag "det" que são os produtos..
'Buscara produto a produto, e o inserirá na nota que esta sendo importada
For Each Det In xDet
xProd = doc.getElementsByTagName("det")(I).XML ' xProd desmembrará o xml pegando produto a produto...
X = Nz(DLookup("IdProd", "tbCadProd", "DescProd = '" & separaEntreDuasStringsXML(Replace(xProd, "'", ""), "<xProd>", "</xProd>") & "'"), 0)
If X <= 0 Then
'Cadastra o Produto, pois ainda nao foi cadastrado
Set rsProd = Db.OpenRecordset("tbCadProd")
rsProd.AddNew
   
    rsProd!DescProd = separaEntreDuasStringsXML(Replace(xProd, "'", ""), "<xProd>", "</xProd>")
    rsProd!Unid = separaEntreDuasStringsXML(xProd, "<uCom>", "</uCom>")
    rsProd!CodFornecedor = separaEntreDuasStringsXML(xProd, "<cProd>", "</cProd>")
    rsProd!NCM = separaEntreDuasStringsXML(xProd, "<NCM>", "</NCM>")
    rsProd!CFOP_ORIGINAL = separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>")
    rsProd!EAN = separaEntreDuasStringsXML(xProd, "<cEANTrib>", "</cEANTrib>")
    'rsProd!Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    
   rsProd!Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
   rsProd!Cd_Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
   
   contProd = contProd + 1
   
   'rsProd!Estoque = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
rsProd.Update
X = Nz(DLookup("IdProd", "tbCadProd", "DescProd = '" & separaEntreDuasStringsXML(Replace(xProd, "'", ""), "<xProd>", "</xProd>") & "'"), 0)
'Insere o produto cadastrado na Nota de compra
Set rsVendaDet = Db.OpenRecordset("tbVendasDet")
rsVendaDet.AddNew
    rsVendaDet!IDVenda = regAtual
    rsVendaDet!IDProd = X
  
    rsVendaDet!Qnt = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    rsVendaDet!ValorUnit = Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",")
    rsVendaDet!ValorTot = Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",") * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    
    rsVendaDet!VlrFrete = Replace(separaEntreDuasStringsXML(xProd, "<vFrete>", "</vFrete>"), ".", ",")
    rsVendaDet!VlrSeguro = Replace(separaEntreDuasStringsXML(xProd, "<vSeg>", "</vSeg>"), ".", ",")
    rsVendaDet!VlrDesc = Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",")
    rsVendaDet!VlrOutro = Replace(separaEntreDuasStringsXML(xProd, "<vOutro>", "</vOutro>"), ".", ",")
    rsVendaDet!CFOP = Replace(separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>"), ".", ",")
    rsVendaDet!Pedido = Replace(separaEntreDuasStringsXML(xProd, "<xPed>", "</xPed>"), ".", ",")
    rsVendaDet!Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    rsVendaDet!Cd_Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    
   
    rsVendaDet!CST = Replace(separaEntreDuasStringsXML(xProd, "<CST>", "</CST>"), ".", ",")
    
    rsVendaDet!BaseCalculo = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!Aliq_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<pICMS>", "</pICMS>"), ".", ",")
    rsVendaDet!Valor_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<vICMS>", "</vICMS>"), ".", ",")
    
    cBaseCalcPis = Replace(separaEntreDuasStringsXML(xProd, "<PISAliq>", "<pPIS>"), ".", ",")
    rsVendaDet!BaseCalc_PisCofins = separaEntreDuasStringsXML(cBaseCalcPis, "<vBC>", "</vBC>")
    
    rsVendaDet!Aliq_PIS = Replace(separaEntreDuasStringsXML(xProd, "<pPIS>", "</pPIS>"), ".", ",")
    rsVendaDet!Valor_PIS = Replace(separaEntreDuasStringsXML(xProd, "<vPIS>", "</vPIS>"), ".", ",")
    rsVendaDet!Aliq_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<pCOFINS>", "</pCOFINS>"), ".", ",")
    rsVendaDet!Valor_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<vCOFINS>", "</vCOFINS>"), ".", ",")
    rsVendaDet!Aliq_IPI = Replace(separaEntreDuasStringsXML(xProd, "<pIPI>", "</pIPI>"), ".", ",")
    rsVendaDet!Valor_IPI = Replace(separaEntreDuasStringsXML(xProd, "<vIPI>", "</vIPI>"), ".", ",")
    rsVendaDet!MVA_ST = Replace(separaEntreDuasStringsXML(xProd, "<pMVAST>", "</pMVAST>"), ".", ",")
    rsVendaDet!Aliq_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<pICMSST>", "</pICMSST>"), ".", ",")
    rsVendaDet!BaseCalc_ST = Replace(separaEntreDuasStringsXML(xProd, "<vBCST>", "</vBCST>"), ".", ",")
    rsVendaDet!Valor_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<vICMSST>", "</vICMSST>"), ".", ",")
    
 
    rsVendaDet!InfoAdicional = Replace(separaEntreDuasStringsXML(xProd, "<infAdProd>", "</infAdProd>"), ".", ",")
    
    rsVendaDet!CustoMedio = DLookup("CMed_Unit", "tbCadProd", "IDprod=" & X) * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    
rsVendaDet.Update

rsVendaDet.Close
rsProd.Close
Set rsVendaDet = Nothing
Set rsProd = Nothing
Else
'Set rsProd = db.OpenRecordset("SELECT * FROM tbCadProd WHERE IdProd = " & x & "")
'rsProd.Edit 'Atualiza o estoque do produto
'    rsProd!Estoque = rsProd!Estoque - Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
' rsProd.Update
'Insere o produto que ja estava cadastrado na Nota de compra
Set rsVendaDet = Db.OpenRecordset("tbVendasDet")
rsVendaDet.AddNew
    rsVendaDet!IDVenda = regAtual
    rsVendaDet!IDProd = X
    
    
    rsVendaDet!Qnt = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    rsVendaDet!ValorUnit = Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",")
    rsVendaDet!ValorTot = Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",") * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    
    'rsVendaDet!VlrFrete = Replace(separaEntreDuasStringsXML(xProd, "<vFrete>", "</vFrete>"), ".", ",")
    'rsVendaDet!VlrSeguro = Replace(separaEntreDuasStringsXML(xProd, "<vSeg>", "</vSeg>"), ".", ",")
    rsVendaDet!VlrDesc = Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",")
    rsVendaDet!VlrOutro = Replace(separaEntreDuasStringsXML(xProd, "<vOutro>", "</vOutro>"), ".", ",")
    rsVendaDet!CFOP = Replace(separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>"), ".", ",")
    rsVendaDet!Pedido = Replace(separaEntreDuasStringsXML(xProd, "<xPed>", "</xPed>"), ".", ",")
    
    rsVendaDet!Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    rsVendaDet!Cd_Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    
    
    
    rsVendaDet!CST = Replace(separaEntreDuasStringsXML(xProd, "<CST>", "</CST>"), ".", ",")
    
    
    rsVendaDet!BaseCalculo = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!Aliq_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<pICMS>", "</pICMS>"), ".", ",")
    rsVendaDet!Valor_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<vICMS>", "</vICMS>"), ".", ",")
    
    cBaseCalcPis = Replace(separaEntreDuasStringsXML(xProd, "<PISAliq>", "<pPIS>"), ".", ",")
    rsVendaDet!BaseCalc_PisCofins = separaEntreDuasStringsXML(cBaseCalcPis, "<vBC>", "</vBC>")
    
    rsVendaDet!Aliq_PIS = Replace(separaEntreDuasStringsXML(xProd, "<pPIS>", "</pPIS>"), ".", ",")
    rsVendaDet!Valor_PIS = Replace(separaEntreDuasStringsXML(xProd, "<vPIS>", "</vPIS>"), ".", ",")
    rsVendaDet!Aliq_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<pCOFINS>", "</pCOFINS>"), ".", ",")
    rsVendaDet!Valor_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<vCOFINS>", "</vCOFINS>"), ".", ",")
    rsVendaDet!Aliq_IPI = Replace(separaEntreDuasStringsXML(xProd, "<pIPI>", "</pIPI>"), ".", ",")
    rsVendaDet!Valor_IPI = Replace(separaEntreDuasStringsXML(xProd, "<vIPI>", "</vIPI>"), ".", ",")
    rsVendaDet!MVA_ST = Replace(separaEntreDuasStringsXML(xProd, "<pMVAST>", "</pMVAST>"), ".", ",")
    rsVendaDet!Aliq_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<pICMSST>", "</pICMSST>"), ".", ",")
    rsVendaDet!BaseCalc_ST = Replace(separaEntreDuasStringsXML(xProd, "<vBCST>", "</vBCST>"), ".", ",")
    rsVendaDet!Valor_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<vICMSST>", "</vICMSST>"), ".", ",")
 
    
    
    
    rsVendaDet!InfoAdicional = Replace(separaEntreDuasStringsXML(xProd, "<infAdProd>", "</infAdProd>"), ".", ",")
    
    rsVendaDet!CustoMedio = DLookup("CMed_Unit", "tbCadProd", "IDProd=" & X) * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
  
    
    
rsVendaDet.Update
'Limpa os dados do recordset e fecha a conecção
rsVendaDet.Close
'rsProd.Close
Set rsVendaDet = Nothing
Set rsProd = Nothing
End If
'Add 1 unidade ao contador para pegar o proximo produto
I = I + 1
Next
Else
'Se abrir xml com erro, será add no formulario "aguarde" o nome dele
Forms!frmaguarde!txt2 = Forms!frmaguarde!txt2 & vbNewLine & NomeArq
Forms!frmaguarde.Requery
End If
'Loop dos arquivos, pega o proximo arquivo
proximoarquivo:

NomeArq = Dir()
Loop


Forms!frmXMLinput.Requery
MsgBox "Vendas-XML Importados! " & vbNewLine & "Notas de Venda: " & contNF_Vend & vbNewLine & "Clientes Novos: " & contCli & vbNewLine & "Produtos: " & contProd, vbInformation, "Sucesso!!!"
Db.Close
Set Db = Nothing
Call DisconnectFromDataBase

End Function


