Attribute VB_Name = "modImportaXML_SAT"
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

Public Function ImportaXML_SAT(LocalXml As String)
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
If doc.validate.errorCode = -1072897500 And (doc.getElementsByTagName("SignatureValue").length) Then
'Set xDet = doc.getElementsByTagName("det")
Set xDet = doc.getElementsByTagName("dest")
'------------------------------------------------------------------------'
'Insere os Dados do Fornecedor, se nao for Cadastrado.
Set rsFor = Db.OpenRecordset("tbcliente")




X = Nz(DLookup("idCliente", "tbcliente", "RazaoSocial = '" & doc.getElementsByTagName("dest/xNome")(0).Text & "'"), 0)     'X buscará o fornecedor na tabela "tbFornecedores"
'x = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & doc.childNodes(2).getElementsByTagName("dest")(0) & "'"), 0) 'X buscará o fornecedor na tabela "tbFornecedores"


                                                    
                                                    
If doc.getElementsByTagName("CNPJ")(1).Text = Forms!frmXMLinput!txt_CNPJ_Empresa Then 'Se CNPJ do emissor for igual ao da empresa, irá cadastrar como venda
Else
GoTo proximoarquivo
End If



If X <= 0 Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o cliente
    rsFor.AddNew
        
        On Error Resume Next
        rsFor!Tipo = "CLIENTE"
        'rsFor!CNPJ = doc.getElementsByTagName("CNPJ_dest")(0).Text
        'rsFor!CNPJ = xDet.Item(0).childNodes(0).Text
        rsFor!CNPJ = doc.getElementsByTagName("dest/xCNPJ")(0).Text
        rsFor!CNPJ = doc.getElementsByTagName("dest/xCPF")(0).Text
        'varCNPJ = xDet.Item(0).childNodes(0).Text
        
        rsFor!RazaoSocial = doc.getElementsByTagName("dest/xNome")(0).Text
        
       
        'Select Case xDet.Item(0).childNodes(3).Text
        'Case "1"
        'rsFor!CRT = "CONTRIBUINTE ICMS"
        'rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        'Case "2"
        'rsFor!CRT = "ISENTO ICMS"
        'rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        'Case "9"
        'rsFor!CRT = "NAO CONTRIBUINTE"
        'Case Else
        'rsFor!CRT = "NAO INFORMADO"
        'End Select
        rsFor!CRT = "CONSUMIDOR"
        
        'rsFor!Email = xDet.Item(0).selectSingleNode("email").Text
        
        
        Set xDet = doc.getElementsByTagName("enderDest")
        
       ' rsFor!Logradouro = xDet.Item(0).selectSingleNode("xLgr").Text
       ' rsFor!Nro = xDet.Item(0).selectSingleNode("nro").Text
       ' rsFor!CEP = xDet.Item(0).selectSingleNode("CEP").Text
'        On Error Resume Next
        
        
       ' rsFor!compl = xDet.Item(0).selectSingleNode("xCpl").Text
       ' rsFor!Bairro = xDet.Item(0).selectSingleNode("xBairro").Text
       ' rsFor!UF = xDet.Item(0).selectSingleNode("UF").Text
       ' rsFor!Municipio = xDet.Item(0).selectSingleNode("xMun").Text
        rsFor!Pais = xDet.Item(0).selectSingleNode("xPais").Text
        'rsFor!fone = xDet.Item(0).selectSingleNode("fone").Text
    
        
        contCli = contCli + 1
        
        
    rsFor.Update
    On Error GoTo 0
'Apos cadastrar o fornecedor, x buscara o ID desse fornecedor para ser utilizado na importação do xml em questao

'x = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & doc.getElementsByTagName("CNPJ_dest")(0).Text & "'"), 0)



X = Nz(DLookup("idCliente", "tbcliente", "RazaoSocial = '" & doc.getElementsByTagName("dest/xNome")(0).Text & "'"), 0)      'X buscará o fornecedor na tabela "tbFornecedores"


Else
'atualiza dados do fornecedor já cadastrado
rsFor.Close
Set rsFor = Db.OpenRecordset("SELECT * FROM tbcliente WHERE idCliente = " & X & "")
rsFor.Edit

        
         On Error Resume Next
        rsFor!Tipo = "CLIENTE"
        
        rsFor!CNPJ = doc.getElementsByTagName("dest/xCNPJ")(0).Text
        rsFor!CNPJ = doc.getElementsByTagName("dest/xCPF")(0).Text
        
        rsFor!RazaoSocial = doc.getElementsByTagName("dest/xNome")(0).Text
        
       
        'Select Case xDet.Item(0).childNodes(3).Text
        'Case "1"
        'rsFor!CRT = "CONTRIBUINTE ICMS"
        'rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        'Case "2"
        'rsFor!CRT = "ISENTO ICMS"
        'rsFor!IE = xDet.Item(0).selectSingleNode("IE").Text
        'Case "9"
        'rsFor!CRT = "NAO CONTRIBUINTE"
        'Case Else
        'rsFor!CRT = "NAO INFORMADO"
        'End Select
        rsFor!CRT = "CONSUMIDOR"
        
        'rsFor!Email = xDet.Item(0).selectSingleNode("email").Text
        
        
        'Set xDet = doc.getElementsByTagName("enderDest")
        
        'rsFor!Logradouro = xDet.Item(0).selectSingleNode("xLgr").Text
        'rsFor!Nro = xDet.Item(0).selectSingleNode("nro").Text
        'rsFor!CEP = xDet.Item(0).selectSingleNode("CEP").Text
'        On Error Resume Next
        
        
        'rsFor!compl = xDet.Item(0).selectSingleNode("xCpl").Text
        'rsFor!Bairro = xDet.Item(0).selectSingleNode("xBairro").Text
        'rsFor!UF = xDet.Item(0).selectSingleNode("UF").Text
        'rsFor!Municipio = xDet.Item(0).selectSingleNode("xMun").Text
        rsFor!Pais = xDet.Item(0).selectSingleNode("xPais").Text
        'rsFor!fone = xDet.Item(0).selectSingleNode("fone").Text
 

rsFor.Update
On Error GoTo 0

End If
rsFor.Close
Set rsFor = Nothing

strSQL = ("update tbCliente set Pais = 'BRASIL', UF = 'SP', cod_Municipio = '3552205', municipio = 'Sorocaba' where Pais is null")
Conn.Execute strSQL

'------------------------------------------------------------------------'
'Dados Principais da Nota de Venda (tbVendas)
Set xDet = doc.getElementsByTagName("det")
Set rsVenda = Db.OpenRecordset("tbVendasSAT")
Set rsVendaPgto = Db.OpenRecordset("tbVendasSATPgto")

'verifica se o xml já foi processado antes pra não duplicar a linha
x1 = Nz(DLookup("ChaveCF", "tbVendasSAT", "ChaveCF = '" & Mid(doc.selectSingleNode("//SignedInfo/Reference").Attributes.getNamedItem("URI").Text, 5, 44) & "'"), 0)
If x1 = Mid(doc.selectSingleNode("//SignedInfo/Reference").Attributes.getNamedItem("URI").Text, 5, 44) Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o fornecedor
GoTo proximoarquivo
Else
End If


rsVenda.AddNew
    rsVenda!IdCliente = X
    rsVenda!NomeConsumidor = doc.getElementsByTagName("dest/xNome")(0).Text
        
    On Error Resume Next
    rsVenda!CPF_CNPJ = doc.getElementsByTagName("dest/xCPF")(0).Text
    On Error GoTo 0
    
    'Necessario essa verificação pois na versao XML 1.10 era somente Data (dEmi) ja na 3.0 mudou para DataHora (dhEmi)
      Dim Day As Integer, month As Integer, year As Integer, hour As Integer, minute As Integer, second As Integer, str As String
      
      str = doc.getElementsByTagName("dEmi")(0).Text
        year = Int(Mid(str, 1, 4))
        month = Int(Mid(str, 5, 2))
        Day = Int(Mid(str, 7, 2))
      str = doc.getElementsByTagName("hEmi")(0).Text
        hour = Int(Mid(str, 1, 2))
        minute = Int(Mid(str, 3, 2))
        second = Int(Mid(str, 5, 2))
        convDate = VBA.DateSerial(year, month, Day) & " " & VBA.TimeSerial(hour, minute, second)
    
    rsVenda!ANO = year
    rsVenda!MES = month
    rsVenda!numCF = doc.getElementsByTagName("nCFe")(0).Text
    rsVenda!ChaveCF = Mid(doc.selectSingleNode("//SignedInfo/Reference").Attributes.getNamedItem("URI").Text, 5, 44)
    
    rsVenda!DataEmissao = convDate
    rsVenda!NumSerieSAT = doc.getElementsByTagName("nserieSAT")(0).Text
    rsVenda!CNPJ_Emissor = doc.getElementsByTagName("emit/CNPJ")(0).Text
    rsVenda!CNPJ_SoftHouse = doc.getElementsByTagName("ide/CNPJ")(0).Text
    
    rsVenda!Vlr_TotalCF = Replace(doc.getElementsByTagName("total/vCFe")(0).Text, ".", ",")
    rsVenda!vDesc = Replace(doc.getElementsByTagName("total/ICMSTot/vDesc")(0).Text, ".", ",")
    rsVenda!vICMS = Replace(doc.getElementsByTagName("total/ICMSTot/vICMS")(0).Text, ".", ",")
    rsVenda!vPIS = Replace(doc.getElementsByTagName("total/ICMSTot/vPIS")(0).Text, ".", ",")
    rsVenda!vCofins = Replace(doc.getElementsByTagName("total/ICMSTot/vCOFINS")(0).Text, ".", ",")
       
    rsVenda!Troco = Replace(doc.getElementsByTagName("pgto/vTroco")(0).Text, ".", ",")
    
    contNF_Vend = contNF_Vend + 1
       
rsVenda.Update
regAtual = Nz(DLookup("idSAT", "tbVendasSAT", "ChaveCF = '" & Mid(doc.selectSingleNode("//SignedInfo/Reference").Attributes.getNamedItem("URI").Text, 5, 44) & "'"), 0)

    Set List = doc.selectNodes("//pgto/MP")
    L = 0
    On Error Resume Next
    For Each List In doc.childNodes
    rsVendaPgto.AddNew
    rsVendaPgto!idSAT = regAtual
    rsVendaPgto!CondPgto = doc.selectNodes("//pgto/MP/cMP")(L).Text
    rsVendaPgto!vPgto = Replace(doc.selectNodes("//pgto/MP/vMP")(L).Text, ".", ",")
    rsVendaPgto.Update
    L = L + 1
    Next List
    On Error GoTo 0

rsVenda.Close
Set rsVenda = Nothing


'------------------------------------------------------------------------'
' Dados dos Produtos

'verifica se o xml já foi processado antes pra não duplicar a linha
X = Nz(DLookup("idSAT", "tbVendasSATDet", "idSAT = " & regAtual & ""), 0)
If X = regAtual Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o cliente
GoTo proximoarquivo
Else
End If


I = 0
xProd = ""
'Aqui é o Loop que percorrerá pela Tag "det" que são os produtos..
'Buscara produ  to a produto, e o inserirá na nota que esta sendo importada
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

Set rsVendaDet = Db.OpenRecordset("tbVendasSATDet")
rsVendaDet.AddNew
    rsVendaDet!idSAT = regAtual
    rsVendaDet!IDProd = X
    
    On Error Resume Next
    
    rsVendaDet!NumItem = doc.getElementsByTagName("det")(I).Attributes.getNamedItem("nItem").Text
    
    rsVendaDet!DescProd = Replace(separaEntreDuasStringsXML(xProd, "<xProd>", "</xProd>"), ".", ",")
    rsVendaDet!NCM = Replace(separaEntreDuasStringsXML(xProd, "<NCM>", "</NCM>"), ".", ",")
    rsVendaDet!UN_Com = Replace(separaEntreDuasStringsXML(xProd, "<uCom>", "</uCom>"), ".", ",")
    rsVendaDet!Det = 0
    rsVendaDet!Texto_Det = 0
  
    rsVendaDet!Qt = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    rsVendaDet!Vlr_Unit = Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",")
    rsVendaDet!Vlr_Prod = Replace(separaEntreDuasStringsXML(xProd, "<vProd>", "</vProd>"), ".", ",") * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    rsVendaDet!Vlr_Desc = Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",")
    rsVendaDet!Vlr_Item = Replace(separaEntreDuasStringsXML(xProd, "<vItem>", "</vItem>"), ".", ",")
    
    rsVendaDet!Texto_Det = Replace(separaEntreDuasStringsXML(xProd, "<xTextoDet>", "</xTextoDet>"), ".", ",")
    
    rsVendaDet!CFOP = Replace(separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>"), ".", ",")
    rsVendaDet!Orig = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
   
    rsVendaDet!CST_ICMS = "0" & Replace(separaEntreDuasStringsXML(xProd, "<CST>", "</CST>"), ".", ",")
    rsVendaDet!bCalcICMS = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!pICMS = Replace(separaEntreDuasStringsXML(xProd, "<pICMS>", "</pICMS>"), ".", ",")
    rsVendaDet!vICMS = Replace(separaEntreDuasStringsXML(xProd, "<vICMS>", "</vICMS>"), ".", ",")
        
    
    rsVendaDet!CST_PIS = Replace(separaEntreDuasStringsXML(xProd, "<PIS><PISAliq><CST>", "</CST>"), ".", ",")
    rsVendaDet!bCalcPIS = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!pPIS = Replace(separaEntreDuasStringsXML(xProd, "<pPIS>", "</pPIS>"), ".", ",")
    rsVendaDet!vPIS = Replace(separaEntreDuasStringsXML(xProd, "<vPIS>", "</vPIS>"), ".", ",")
    
    rsVendaDet!CST_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<COFINS><COFINSAliq><CST>", "</CST>"), ".", ",")
    rsVendaDet!bCalcCofins = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!pCofins = Replace(separaEntreDuasStringsXML(xProd, "<pCOFINS>", "</pCOFINS>"), ".", ",")
    rsVendaDet!vCofins = Replace(separaEntreDuasStringsXML(xProd, "<vCOFINS>", "</vCOFINS>"), ".", ",")
    On Error GoTo 0
    
    If rsVendaDet!CST_ICMS = "060" Then
    rsVendaDet!CST_PIS = "04"
    rsVendaDet!CST_Cofins = "04"
    Else
    End If
    'rsVendaDet!InfoAdicional = Replace(separaEntreDuasStringsXML(xProd, "<infAdProd>", "</infAdProd>"), ".", ",")
    
    'rsVendaDet!CustoMedio = DLookup("CMed_Unit", "tbCadProd", "IDprod=" & X) * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    
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
Set rsVendaDet = Db.OpenRecordset("tbVendasSATDet")
rsVendaDet.AddNew
    rsVendaDet!idSAT = regAtual
    rsVendaDet!IDProd = X
    
    rsVendaDet!DescProd = Replace(separaEntreDuasStringsXML(xProd, "<xProd>", "</xProd>"), ".", ",")
    rsVendaDet!NCM = Replace(separaEntreDuasStringsXML(xProd, "<NCM>", "</NCM>"), ".", ",")
    rsVendaDet!UN_Com = Replace(separaEntreDuasStringsXML(xProd, "<uCom>", "</uCom>"), ".", ",")
        
    
    rsVendaDet!NumItem = doc.getElementsByTagName("det")(I).Attributes.getNamedItem("nItem").Text
    
    On Error Resume Next
    rsVendaDet!Qt = Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    rsVendaDet!Vlr_Unit = Replace(separaEntreDuasStringsXML(xProd, "<vUnCom>", "</vUnCom>"), ".", ",")
    rsVendaDet!Vlr_Prod = Replace(separaEntreDuasStringsXML(xProd, "<vProd>", "</vProd>"), ".", ",") * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    rsVendaDet!Vlr_Desc = Replace(separaEntreDuasStringsXML(xProd, "<vDesc>", "</vDesc>"), ".", ",")
    rsVendaDet!Vlr_Item = Replace(separaEntreDuasStringsXML(xProd, "<vItem>", "</vItem>"), ".", ",")
    
    rsVendaDet!Texto_Det = Replace(separaEntreDuasStringsXML(xProd, "<xTextoDet>", "</xTextoDet>"), ".", ",")
    
    rsVendaDet!CFOP = Replace(separaEntreDuasStringsXML(xProd, "<CFOP>", "</CFOP>"), ".", ",")
    rsVendaDet!Orig = separaEntreDuasStringsXML(xProd, "<orig>", "</orig>")
   
    rsVendaDet!CST_ICMS = "0" & Replace(separaEntreDuasStringsXML(xProd, "<CST>", "</CST>"), ".", ",")
    rsVendaDet!bCalcICMS = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!pICMS = Replace(separaEntreDuasStringsXML(xProd, "<pICMS>", "</pICMS>"), ".", ",")
    rsVendaDet!vICMS = Replace(separaEntreDuasStringsXML(xProd, "<vICMS>", "</vICMS>"), ".", ",")
    
   
    rsVendaDet!CST_PIS = Replace(separaEntreDuasStringsXML(xProd, "<PIS><PISAliq><CST>", "</CST>"), ".", ",")
    rsVendaDet!bCalcPIS = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!pPIS = Replace(separaEntreDuasStringsXML(xProd, "<pPIS>", "</pPIS>"), ".", ",")
    rsVendaDet!vPIS = Replace(separaEntreDuasStringsXML(xProd, "<vPIS>", "</vPIS>"), ".", ",")
    
    rsVendaDet!CST_Cofins = Replace(separaEntreDuasStringsXML(xProd, "<COFINS><COFINSAliq><CST>", "</CST>"), ".", ",")
    rsVendaDet!bCalcCofins = Replace(separaEntreDuasStringsXML(xProd, "<vBC>", "</vBC>"), ".", ",")
    rsVendaDet!pCofins = Replace(separaEntreDuasStringsXML(xProd, "<pCOFINS>", "</pCOFINS>"), ".", ",")
    rsVendaDet!vCofins = Replace(separaEntreDuasStringsXML(xProd, "<vCOFINS>", "</vCOFINS>"), ".", ",")
    On Error GoTo 0
    'rsVendaDet!CustoMedio = DLookup("CMed_Unit", "tbCadProd", "IDProd=" & X) * Replace(separaEntreDuasStringsXML(xProd, "<qCom>", "</qCom>"), ".", ",")
    
    If rsVendaDet!CST_ICMS = "060" Then
    rsVendaDet!CST_PIS = "04"
    rsVendaDet!CST_Cofins = "04"
    Else
    End If
    
    
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


'PROCESSA AS VENDAS SAT COMO UMA VENDA NA TB VENDAS
'header
Call ConnectToDataBase

'Delete q1
'FROM tbVendasDet as q1
'INNER Join
'tbVendas As q2
'ON q1.IDVenda = q2.ID
'INNER Join
'tbVendasSAT As q3
'on q2.ChaveNF = q3.ChaveCF;

'Delete q1
'FROM tbVendas as q1
'INNER Join
'tbVendasSAT As q3
'on q1.ChaveNF = q3.ChaveCF;

strSQL = ("update tbVendasSatDet set Vlr_Desc = 0 where Vlr_Desc is null")
Conn.Execute strSQL

strSQL = ("update tbVendassatdet set pPIS = pPIS * 100 where pPIS <1;")
Conn.Execute strSQL

strSQL = ("update tbVendassatdet set pCofins = pCofins * 100 where pCofins <1;")
Conn.Execute strSQL

'CALCULA IPI DO SAT
strSQL = ("update tbvendassatdet as q1 inner join tbcadprod as q2 on q1.IdProd = q2.IDProd set q1.bCalcIPI = q1.bCalcICMS, q1.vIPI = q1.bCalcICMS * 0.036 where PROD_FINAL = 'SIM' and vIPI = 0;")
Conn.Execute strSQL
'CALCULA IPI DO SAT
strSQL = ("update tbvendassat as q1 inner join (select idSAT, sum(vIPI) as vIPI from tbvendassatdet group by idSAT) as q2 on q1.idSAT = q2.idSAT Set q1.vIPI = q2.vIPI where q1.vIPI = 0;")
Conn.Execute strSQL


strSQL = ("insert into tbVendas(ANO, MES,`Status`, IdCliente, TipoNF, NatOperacao,                             DataEmissao, NumNF, Serie, ChaveNF, VlrTotalProdutos, VlrTotalFrete, VlrTotalSeguro, VlrDesconto, VlrDespesas, ICMS_BaseCalc, ICMS_Valor, ICMS_ST_BaseCalc, ICMS_ST_Valor, IPI_Valor, PIS_valor, COFINS_Valor, VlrTOTALNF) " & _
"select               q1.ANO, q1.MES, 'ATIVO', q1.idCliente,'1-SAIDA','Venda Cupom Fiscal SAT', q1.DataEmissao, q1.numCF, '1',   q1.ChaveCF, q1.Vlr_TotalCF,      0,             0,              q1.vDesc,       0,          q1.Vlr_TotalCF,   q1.vICMS,      0,               0,             0,          q1.vPIS,      q1.vCofins,  q1.Vlr_TotalCF       from tbVendasSAT as q1 left outer join tbVendas as q2 on q1.ChaveCF = q2.ChaveNF where q2.ChaveNF is null; ")
Conn.Execute strSQL

'line
strSQL = ("insert into tbVendasDet (IDVenda, IDProd,    Qnt,   ValorUnit,      ValorTot, VlrDesc, VlrOutro, CFOP,    Cd_Origem, CST,        BaseCalculo, Aliq_ICMS, Valor_ICMS, BaseCalc_PisCofins, Aliq_PIS, Valor_PIS, Aliq_Cofins, Valor_Cofins, Aliq_IPI, Valor_IPI, MVA_ST, Aliq_ICMS_ST, BaseCalc_ST, Valor_ICMS_ST, DAS_Aliq, DAS_Valor, CST_ICMS, CST_PIS, CST_Cofins) " & _
"select                   q1.ID,   q3.IdProd, q3.Qt, q3.Vlr_Unit, q3.Vlr_Item, 0,       0,        q3.CFOP, q3.Orig,    q3.CST_ICMS, q3.bCalcICMS, q3.pICMS, q3.vICMS, q3.bCalcPIS,         q3.pPIS, q3.vPIS,     q3.pCofins, q3.vCofins,    0,        0,         0,      0,            0,           0            , 0       , 0        , q3.CST_ICMS, q3.CST_PIS, q3.CST_Cofins " & _
"from tbVendas as q1 inner join tbVendasSAT as q2 on q1.ChaveNF = q2.ChaveCF left outer join tbVendasDet as q4 on q1.ID = q4.IDVenda left outer join tbVendasSATDet as q3 on q2.idSAT = q3.idSAT where q4.IDVenda is null;")
Conn.Execute strSQL
'PROCESSA AS VENDAS SAT COMO UMA VENDA NA TB VENDAS



Forms!frmXMLinput.Requery
MsgBox "Vendas-XML Importados! " & vbNewLine & "Notas de Venda: " & contNF_Vend & vbNewLine & "Clientes Novos: " & contCli & vbNewLine & "Produtos: " & contProd, vbInformation, "Sucesso!!!"
Db.Close
Set Db = Nothing
Call DisconnectFromDataBase





End Function

