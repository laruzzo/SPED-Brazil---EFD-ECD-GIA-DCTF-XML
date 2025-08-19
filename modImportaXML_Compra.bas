Attribute VB_Name = "modImportaXML_Compra"
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
Public Function ImportaXML_Compra(LocalXml As String)
'---------------------------------------------------------------'
'                Criado por FabioPaes                           '
'           Em 12/02/2017 para MAXIMOACCESS                     '
' Em caso de correçoes reportar a origem para atualizar o codigo'
'---------------------------------------------------------------'

Call ConnectToDataBase



Dim doc As DOMDocument
Dim xDet As IXMLDOMNodeList
Dim NomeArq As String
Dim Db As Database
Dim rsFor, rsProd, rsCompra, rsCompDet As DAO.Recordset
Dim xProd As String
Dim I, X, regAtual As Integer

Dim contNF_Ent As Integer
Dim contNF_Vend As Integer
Dim contProd As Integer
Dim contFor As Integer
Dim contCli As Integer

contNF_Ent = 0
contProd = 0
contFor = 0


DoCmd.OpenForm "frmAguarde"
Forms!frmaguarde!txt2 = "Xml Com Erros:"
Diret = LocalXml
NomeArq = Dir(Diret & "*.XML", vbArchive)

Set Db = CurrentDb()
Set doc = New DOMDocument

Dim cSaldoQt As Integer
Dim cCustoMedio As Double
Dim cCustoAtual As Double
Dim cEntradaQt As Integer
Dim cCustoEntrada As Double
Dim cCustoTotEntrada As Double

Dim cBaseCalcPis As String

'Buscará todos os arquivos com extenção .xml da pasta selecionada
Do While NomeArq <> ""
doc.Load (LocalXml & NomeArq) 'Pega a Pasta e o Nome do primeiro arquivo....
'Verifica se o Arquivo foi aberto corretamente e se possui chave. Se possuir importa, se nao Pula pra o Proximo!
If doc.validate.errorCode = -1072897500 And (doc.getElementsByTagName("chNFe").length) Then
Set xDet = doc.getElementsByTagName("det")
'------------------------------------------------------------------------'
'Insere os Dados do Fornecedor, se nao for Cadastrado.
Set rsFor = Db.OpenRecordset("tbfornecedor")
X = Nz(DLookup("IdFor", "tbfornecedor", "Cnpj = '" & doc.getElementsByTagName("CNPJ")(0).Text & "'"), 0) 'X buscará o fornecedor na tabela "tbfornecedor"

If doc.getElementsByTagName("CNPJ")(0).Text = Forms!frmXMLinput!txt_CNPJ_Empresa Then 'Se CNPJ for igual ao da empresa, não irá cadastrar como compra.
GoTo proximoarquivo
Else
End If



If X <= 0 Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o fornecedor
    rsFor.AddNew
'     On Error Resume Next
        
        'rsFor!NomeForn = doc.getElementsByTagName("xFant")(0).Text
        rsFor!Tipo = "FORNECEDOR"
        rsFor!CNPJ = doc.getElementsByTagName("CNPJ")(0).Text
        
        rsFor!RazaoSocial = doc.getElementsByTagName("xNome")(0).Text
        rsFor!IE = doc.getElementsByTagName("IE")(0).Text
       
        Select Case doc.getElementsByTagName("CRT")(0).Text
        Case "1"
        rsFor!CRT = "SIMPLES NACIONAL"
        Case "2"
        rsFor!CRT = "SIMPLES NACIONAL"
        Case "3"
        rsFor!CRT = "REGIME NORMAL"
        Case Else
        rsFor!CRT = "NAO INFORMADO"
        End Select
         
        On Error Resume Next
        rsFor!Logradouro = doc.getElementsByTagName("xLgr")(0).Text
        rsFor!Nro = doc.getElementsByTagName("nro")(0).Text
        rsFor!CEP = doc.getElementsByTagName("CEP")(0).Text
        rsFor!compl = doc.getElementsByTagName("xCpl")(0).Text
        rsFor!Bairro = doc.getElementsByTagName("xBairro")(0).Text
        rsFor!UF = doc.getElementsByTagName("UF")(0).Text
        rsFor!Municipio = doc.getElementsByTagName("xMun")(0).Text
        rsFor!Pais = doc.getElementsByTagName("xPais")(0).Text
        rsFor!Fone = doc.getElementsByTagName("fone")(0).Text
        rsFor!Email = doc.getElementsByTagName("Email")(0).Text
        On Error GoTo -1
        
        contFor = contFor + 1
        
    rsFor.Update
'Apos cadastrar o fornecedor, x buscara o ID desse fornecedor para ser utilizado na importação do xml em questao
X = Nz(DLookup("IdFor", "tbfornecedor", "Cnpj = '" & doc.getElementsByTagName("CNPJ")(0).Text & "'"), 0)

Else
'atualiza dados do fornecedor já cadastrado
rsFor.Close
Set rsFor = Db.OpenRecordset("SELECT * FROM tbfornecedor WHERE IdFor = " & X & "")
rsFor.Edit
'On Error Resume Next
        rsFor!Tipo = "FORNECEDOR"
        'rsFor!NomeForn = doc.getElementsByTagName("xFant")(0).Text
        'rsFor!CNPJ = doc.getElementsByTagName("CNPJ")(0).Text
        rsFor!RazaoSocial = doc.getElementsByTagName("xNome")(0).Text
        rsFor!IE = doc.getElementsByTagName("IE")(0).Text
        
        Select Case doc.getElementsByTagName("CRT")(0).Text
        Case "1"
        rsFor!CRT = "SIMPLES NACIONAL"
        Case "2"
        rsFor!CRT = "SIMPLES NACIONAL"
        Case "3"
        rsFor!CRT = "REGIME NORMAL"
        Case Else
        rsFor!CRT = "NAO INFORMADO"
        End Select
        
        rsFor!Logradouro = doc.getElementsByTagName("xLgr")(0).Text
        rsFor!Nro = doc.getElementsByTagName("nro")(0).Text
        rsFor!CEP = doc.getElementsByTagName("CEP")(0).Text
        
        rsFor!Bairro = doc.getElementsByTagName("xBairro")(0).Text
        
        On Error Resume Next
        rsFor!compl = doc.getElementsByTagName("xCpl")(0).Text
        On Error GoTo -1
        
        rsFor!UF = doc.getElementsByTagName("UF")(0).Text
        rsFor!Municipio = doc.getElementsByTagName("xMun")(0).Text
        rsFor!Pais = doc.getElementsByTagName("xPais")(0).Text
        
        On Error Resume Next
        rsFor!Fone = doc.getElementsByTagName("fone")(0).Text
        rsFor!Email = doc.getElementsByTagName("Email")(0).Text
        On Error GoTo -1
        
rsFor.Update

'On Error GoTo 0

End If
rsFor.Close
Set rsFor = Nothing

'------------------------------------------------------------------------'
'Dados Principais da Nota de Compra (tbCompras)
Set rsCompra = Db.OpenRecordset("tbCompras")

'verifica se o xml já foi processado antes pra não duplicar a linha
x1 = Nz(DLookup("ChaveNF", "tbCompras", "ChaveNF = '" & doc.getElementsByTagName("chNFe")(0).Text & "'"), 0)
If x1 = doc.getElementsByTagName("chNFe")(0).Text Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o fornecedor
GoTo proximoarquivo
Else
End If


rsCompra.AddNew
    rsCompra!IdFornecedor = X
    'Necessario essa verificação pois na versao XML 1.10 era somente Data (dEmi) ja na 3.0 mudou para DataHora (dhEmi)
    If (doc.getElementsByTagName("dhEmi").length) Then
    rsCompra!DataEmissao = Format(Left(doc.getElementsByTagName("dhEmi")(0).Text, 10), "dd/mm/yyyy")
    Else
    rsCompra!DataEmissao = Format(doc.getElementsByTagName("dEmi")(0).Text, "dd/mm/yyyy")
    End If
        'Passa os totais da NFe para a variavel xProd
        'xProd = doc.getElementsByTagName("total")(0).XML
        'Valor Bruto sem desconto
        'rsCompra!ValorNFB = Replace(separaEntreDuasStringsXML(xProd, "<vProd>", "</vProd>"), ".", ",")
        'Valor Liquido "Valor Bruto-descontos"
        'rsCompra!ValorNFL = Replace(separaEntreDuasStringsXML(xProd, "<vNF>", "</vNF>"), ".", ",")
    
    
    Select Case doc.getElementsByTagName("tpNF")(0).Text
    Case 0
    rsCompra!TipoNF = "0-ENTRADA"
    Case 1
    rsCompra!TipoNF = "1-SAIDA"
    End Select
    
    
    rsCompra!NumNF = doc.getElementsByTagName("nNF")(0).Text
    rsCompra!Serie = doc.getElementsByTagName("serie")(0).Text
    rsCompra!chavenf = doc.getElementsByTagName("chNFe")(0).Text
    rsCompra!NatOperacao = doc.getElementsByTagName("natOp")(0).Text
   'rsCompra!ConsumidorFinal = doc.getElementsByTagName("indFinal")(0).Text
   'rsCompra!DestOperacao = doc.getElementsByTagName("idDest")(0).Text
    
    Set xtotal = doc.getElementsByTagName("ICMSTot")
    rsCompra!VlrTotalProdutos = Replace(xtotal.Item(0).selectSingleNode("vProd").Text, ".", ",")
    rsCompra!VlrTotalFrete = Replace(xtotal.Item(0).selectSingleNode("vFrete").Text, ".", ",")
    rsCompra!VlrTotalSeguro = Replace(xtotal.Item(0).selectSingleNode("vSeg").Text, ".", ",")
    rsCompra!VlrDesconto = Replace(xtotal.Item(0).selectSingleNode("vDesc").Text, ".", ",")
    rsCompra!VlrDespesas = Replace(xtotal.Item(0).selectSingleNode("vOutro").Text, ".", ",")
    rsCompra!ICMS_BaseCalc = Replace(xtotal.Item(0).selectSingleNode("vBC").Text, ".", ",")
    rsCompra!ICMS_Valor = Replace(xtotal.Item(0).selectSingleNode("vICMS").Text, ".", ",")
    rsCompra!ICMS_ST_BaseCalc = Replace(xtotal.Item(0).selectSingleNode("vBCST").Text, ".", ",")
    rsCompra!ICMS_ST_Valor = Replace(xtotal.Item(0).selectSingleNode("vST").Text, ".", ",")
    
  
    
    rsCompra!IPI_Valor = Replace(xtotal.Item(0).selectSingleNode("vIPI").Text, ".", ",")
    rsCompra!PIS_Valor = Replace(xtotal.Item(0).selectSingleNode("vPIS").Text, ".", ",")
    rsCompra!COFINS_Valor = Replace(xtotal.Item(0).selectSingleNode("vCOFINS").Text, ".", ",")
    rsCompra!VlrTOTALNF = Replace(xtotal.Item(0).selectSingleNode("vNF").Text, ".", ",")
    xtotal = ""
    
    contNF_Ent = contNF_Ent + 1
    
rsCompra.Update
regAtual = Nz(DLookup("ID", "tbCompras", "ChaveNF = '" & doc.getElementsByTagName("chNFe")(0).Text & "'"), 0)



rsCompra.Close
Set rsCompra = Nothing

'Registra no contas a pagar
strSQL = ("INSERT INTO tb_Detalhe_Boletos_Compras ( DtEmissao, Id_Fornecedor, Fornecedor, ValorOriginal, chave_NFe, STATUS, NumBoleto ) " & _
                "SELECT tbCompras.DataEmissao, tbFornecedor.IDFor, tbFornecedor.RazaoSocial, tbCompras.VlrTOTALNF, tbCompras.ChaveNF, 'ABERTO' AS STATUS, 'LANC AUTOMATICO' AS boleto " & _
                "FROM tbFornecedor INNER JOIN tbCompras on tbFornecedor.IDFor = tbCompras.IdFornecedor WHERE tbCompras.ChaveNF='" & doc.getElementsByTagName("chNFe")(0).Text & "';")


'Call ConnectToDataBase
Conn.Execute strSQL
'Registra no contas a pagar

'------------------------------------------------------------------------'
' Dados dos Produtos

'verifica se o xml já foi processado antes pra não duplicar a linha
X = Nz(DLookup("IDCompra", "tbComprasDet", "IDCompra = " & regAtual & ""), 0)
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
X = Nz(DLookup("IdProd", "tbCadProd", "DescProd = '" & separaEntreDuasStringsXML(Replace(Replace(xProd, "'", ""), Chr(34), ""), "<xProd>", "</xProd>") & "'"), 0)
If X <= 0 Then
'Cadastra o Produto, pois ainda nao foi cadastrado
Set rsProd = Db.OpenRecordset("tbCadProd")
rsProd.AddNew
    rsProd!DescProd = separaEntreDuasStringsXML(Replace(Replace(xProd, "'", ""), Chr(34), ""), "<xProd>", "</xProd>")
    rsProd!Unid = separaEntreDuasStringsXML(xProd, "<uCom>", "</uCom>")
    rsProd!CodFornecedor = separaEntreDuasStringsXML(xProd, "<cProd>", "</cProd>")
    rsProd!NCM = separaEntreDuasStringsXML(xProd, "<NCM>", "</NCM>")
    rsProd!CFOP_ORIGINAL = separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>")
    rsProd!EAN = separaEntreDuasStringsXML(xProd, "<cEANTrib>", "</cEANTrib>")
    
    rsProd!Cd_Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    rsProd!Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
        
    'rsProd!Estoque = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    'estoque será tratado na procedure a parte
    'rsProd!CustoMedio = ((Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",") * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")) + Replace(separaEntreDuasStringsXML(xProd, "<vFrete>", "</vFrete>"), ".", ",") + Replace(separaEntreDuasStringsXML(xProd, "<vSeg>", "</vSeg>"), ".", ",") + Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",") + Replace(separaEntreDuasStringsXML(xProd, "<vOutro>", "</vOutro>"), ".", ",")) / Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
        
    contProd = contProd + 1
    
rsProd.Update
X = Nz(DLookup("IDProd", "tbCadProd", "DescProd = '" & separaEntreDuasStringsXML(Replace(Replace(xProd, "'", ""), Chr(34), ""), "<xProd>", "</xProd>") & "'"), 0)
Z = separaEntreDuasStringsXML(Replace(Replace(xProd, " '", ""), Chr(34), ""), "<xProd>", "</xProd>") & ""
'Insere o produto cadastrado na Nota de compra
Set rsCompDet = Db.OpenRecordset("tbComprasDet")
rsCompDet.AddNew
    rsCompDet!IDCompra = regAtual
    rsCompDet!IDProd = X
    rsCompDet!Qnt = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    rsCompDet!ValorUnit = Round(Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ","), 2)
    rsCompDet!ValorTot = Round(Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",") * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ","), 2)
    
    rsCompDet!VlrFrete = Replace(separaEntreDuasStringsXML(xProd, "<vFrete>", "</vFrete>"), ".", ",")
    rsCompDet!VlrSeguro = Replace(separaEntreDuasStringsXML(xProd, "<vSeg>", "</vSeg>"), ".", ",")
    rsCompDet!VlrDesc = Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",")
    rsCompDet!VlrOutro = Replace(separaEntreDuasStringsXML(xProd, "<vOutro>", "</vOutro>"), ".", ",")
    rsCompDet!CFOP = Replace(separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>"), ".", ",")
    rsCompDet!Pedido = Replace(separaEntreDuasStringsXML(xProd, "<xPed>", "</xPed>"), ".", ",")
    
    rsCompDet!Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    rsCompDet!Cd_Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
   
    rsCompDet!CST = Replace(separaEntreDuasStringsXML(xProd, "<CST>", "</CST>"), ".", ",")
    
    rsCompDet!BaseCalculo = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsCompDet!Aliq_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<pICMS>", "</pICMS>"), ".", ",")
    rsCompDet!Valor_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<vICMS>", "</vICMS>"), ".", ",")
    
    'em discussão com o STF
    cBaseCalcPis = Replace(separaEntreDuasStringsXML(xProd, "<PISAliq>", "<pPIS>"), ".", ",")
    rsCompDet!BaseCalc_PisCofins = separaEntreDuasStringsXML(cBaseCalcPis, "<vBC>", "</vBC>")
    
    rsCompDet!Aliq_PIS = Replace(separaEntreDuasStringsXML(xProd, "<pPIS>", "</pPIS>"), ".", ",")
    rsCompDet!Valor_PIS = Replace(separaEntreDuasStringsXML(xProd, "<vPIS>", "</vPIS>"), ".", ",")
    rsCompDet!Aliq_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<pCOFINS>", "</pCOFINS>"), ".", ",")
    rsCompDet!Valor_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<vCOFINS>", "</vCOFINS>"), ".", ",")
    rsCompDet!Aliq_IPI = Replace(separaEntreDuasStringsXML(xProd, "<pIPI>", "</pIPI>"), ".", ",")
    rsCompDet!Valor_IPI = Replace(separaEntreDuasStringsXML(xProd, "<vIPI>", "</vIPI>"), ".", ",")
    rsCompDet!MVA_ST = Replace(separaEntreDuasStringsXML(xProd, "<pMVAST>", "</pMVAST>"), ".", ",")
    rsCompDet!Aliq_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<pICMSST>", "</pICMSST>"), ".", ",")
    rsCompDet!BaseCalc_ST = Replace(separaEntreDuasStringsXML(xProd, "<vBCST>", "</vBCST>"), ".", ",")
    rsCompDet!Valor_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<vICMSST>", "</vICMSST>"), ".", ",")
    
    
    rsCompDet!InfoAdicional = Replace(separaEntreDuasStringsXML(xProd, "<infAdProd>", "</infAdProd>"), ".", ",")
    
    
rsCompDet.Update

rsCompDet.Close
rsProd.Close
Set rsCompDet = Nothing
Set rsProd = Nothing
Else
'Set rsProd = db.OpenRecordset("SELECT * FROM tbCadProd WHERE IdProd = " & x & "")
'rsProd.Edit 'Atualiza o estoque do produto
'    rsProd!Estoque = rsProd!Estoque + Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    'custo medio ponderado movel
'    cSaldoQt = DLookup("SELECT Estoque from tbCadProd where IDProd = " & x & "", acViewNormal, acReadOnly)
'    cCustoMedio = DLookup("SELECT CustoMedio from tbCadProd where IDProd = " & x & "", acViewNormal, acReadOnly)
'    cCustoAtual = cSaldoQt * cCustoMedio
'    cEntradaQt = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
'    cCustoEntrada = ((Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",") * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")) + Replace(separaEntreDuasStringsXML(xProd, "<vFrete>", "</vFrete>"), ".", ",") + Replace(separaEntreDuasStringsXML(xProd, "<vSeg>", "</vSeg>"), ".", ",") - Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",") + Replace(separaEntreDuasStringsXML(xProd, "<vOutro>", "</vOutro>"), ".", ",")) / Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
'    cCustoTotEntrada = cEntradaQt * cCustoEntrada
'    rsProd!CustoMedio = (cCustoAtual + cCustoTotEntrada) / (cSaldoQt + cEntradaQt)
'    cSaldoQt = 0
'    cCustoMedio = 0
'    cCustoAtual = 0
'    cEntradaQt = 0
'    cCustoEntrada = 0
'    cCustoTotEntrada = 0
    'custo medio ponderado movel
' rsProd.Update
'Insere o produto que ja estava cadastrado na Nota de compra
Set rsCompDet = Db.OpenRecordset("tbComprasDet")
rsCompDet.AddNew
    rsCompDet!IDCompra = regAtual
    rsCompDet!IDProd = X
   
   
   
    rsCompDet!Qnt = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    
    
    rsCompDet!VlrFrete = Replace(separaEntreDuasStringsXML(xProd, "<vFrete>", "</vFrete>"), ".", ",")
       
    
       
    rsCompDet!ValorUnit = Round(Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ","), 2)
    rsCompDet!ValorTot = Round(Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ","), 2) * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
        
    rsCompDet!VlrSeguro = Replace(separaEntreDuasStringsXML(xProd, "<vSeg>", "</vSeg>"), ".", ",")
       
    rsCompDet!VlrDesc = Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",")
    rsCompDet!VlrOutro = Replace(separaEntreDuasStringsXML(xProd, "<vOutro>", "</vOutro>"), ".", ",")
    rsCompDet!CFOP = Replace(separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>"), ".", ",")
    rsCompDet!Pedido = Replace(separaEntreDuasStringsXML(xProd, "<xPed>", "</xPed>"), ".", ",")
    
    rsCompDet!Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    rsCompDet!Cd_Origem = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
    
    rsCompDet!CST = Replace(separaEntreDuasStringsXML(xProd, "<CST>", "</CST>"), ".", ",")
    
    rsCompDet!BaseCalculo = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsCompDet!Aliq_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<pICMS>", "</pICMS>"), ".", ",")
    rsCompDet!Valor_ICMS = Replace(separaEntreDuasStringsXML(xProd, "<vICMS>", "</vICMS>"), ".", ",")
    
    'Discussão do STF
    cBaseCalcPis = Replace(separaEntreDuasStringsXML(xProd, "<PISAliq>", "<pPIS>"), ".", ",")
    rsCompDet!BaseCalc_PisCofins = separaEntreDuasStringsXML(cBaseCalcPis, "<vBC>", "</vBC>")
    
    rsCompDet!Aliq_PIS = Replace(separaEntreDuasStringsXML(xProd, "<pPIS>", "</pPIS>"), ".", ",")
    rsCompDet!Valor_PIS = Replace(separaEntreDuasStringsXML(xProd, "<vPIS>", "</vPIS>"), ".", ",")
    rsCompDet!Aliq_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<pCOFINS>", "</pCOFINS>"), ".", ",")
    rsCompDet!Valor_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<vCOFINS>", "</vCOFINS>"), ".", ",")
    rsCompDet!Aliq_IPI = Replace(separaEntreDuasStringsXML(xProd, "<pIPI>", "</pIPI>"), ".", ",")
    rsCompDet!Valor_IPI = Replace(separaEntreDuasStringsXML(xProd, "<vIPI>", "</vIPI>"), ".", ",")
    rsCompDet!MVA_ST = Replace(separaEntreDuasStringsXML(xProd, "<pMVAST>", "</pMVAST>"), ".", ",")
    rsCompDet!Aliq_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<pICMSST>", "</pICMSST>"), ".", ",")
    rsCompDet!BaseCalc_ST = Replace(separaEntreDuasStringsXML(xProd, "<vBCST>", "</vBCST>"), ".", ",")
    rsCompDet!Valor_ICMS_ST = Replace(separaEntreDuasStringsXML(xProd, "<vICMSST>", "</vICMSST>"), ".", ",")
    
    
    rsCompDet!InfoAdicional = Replace(separaEntreDuasStringsXML(xProd, "<infAdProd>", "</infAdProd>"), ".", ",")
   
   
rsCompDet.Update
'Limpa os dados do recordset e fecha a conexão
rsCompDet.Close
'rsProd.Close
Set rsCompDet = Nothing
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
MsgBox "Compras-XML Importados! " & vbNewLine & "Notas de Entrada: " & contNF_Ent & vbNewLine & "Produtos Novos: " & contProd & vbNewLine & "Fornecedores Novos: " & contFor, vbInformation, "Sucesso!!!"
Db.Close
Set Db = Nothing
Call DisconnectFromDataBase

End Function

