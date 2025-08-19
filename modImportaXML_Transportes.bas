Attribute VB_Name = "modImportaXML_Transportes"
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

Public Function ImportaXML_Transportes(LocalXml As String)
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

Set Db = CurrentDb()
Set doc = New DOMDocument

'Buscará todos os arquivos com extenção .xml da pasta selecionada
Do While NomeArq <> ""
doc.Load (LocalXml & NomeArq) 'Pega a Pasta e o Nome do primeiro arquivo....
'Verifica se o Arquivo foi aberto corretamente e se possui chave. Se possuir importa, se nao Pula pra o Proximo!
If doc.validate.errorCode = -1072897500 And (doc.getElementsByTagName("chCTe").length) Then
'Set xDet = doc.getElementsByTagName("det")

Set xide = doc.getElementsByTagName("ide")

    
Set xEmit = doc.getElementsByTagName("emit")
Set xEmitEnd = doc.getElementsByTagName("enderEmit")
Set xRem = doc.getElementsByTagName("rem")
Set xRemEnd = doc.getElementsByTagName("enderReme")
Set xDest = doc.getElementsByTagName("dest")
Set xDestEnd = doc.getElementsByTagName("enderDest")


Set xDet = doc.getElementsByTagName("infNFe")


'------------------------------------------------------------------------'
'Insere os Dados do Fornecedor, se nao for Cadastrado.
Set rsFor = Db.OpenRecordset("tbFornecedor")




X = Nz(DLookup("IdFor", "tbFornecedor", "Cnpj = '" & xEmit.Item(0).selectSingleNode("CNPJ").Text & "'"), 0)     'X buscará o fornecedor na tabela "tbFornecedores"
'x = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & doc.childNodes(2).getElementsByTagName("dest")(0) & "'"), 0) 'X buscará o fornecedor na tabela "tbFornecedores"


If X <= 0 Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar a transportadora
    rsFor.AddNew
        
        On Error Resume Next
        rsFor!Tipo = "TRANSPORTADORA"
        rsFor!CNPJ = xEmit.Item(0).selectSingleNode("CNPJ").Text
        varCNPJ = xEmit.Item(0).selectSingleNode("CNPJ").Text
        
        rsFor!RazaoSocial = xEmit.Item(0).selectSingleNode("xNome").Text
        rsFor!IE = xEmit.Item(0).selectSingleNode("IE").Text
       
        rsFor!CRT = "NAO INFORMADO"
                
        rsFor!Logradouro = xEmitEnd.Item(0).selectSingleNode("xLgr").Text
        rsFor!Nro = xEmitEnd.Item(0).selectSingleNode("nro").Text
        rsFor!CEP = xEmitEnd.Item(0).selectSingleNode("CEP").Text
        
        
        'rsFor!Compl = xDet.Item(0).selectSingleNode("xCpl").Text
        rsFor!Bairro = xEmitEnd.Item(0).selectSingleNode("xBairro").Text
        rsFor!UF = xEmitEnd.Item(0).selectSingleNode("UF").Text
        rsFor!Municipio = xEmitEnd.Item(0).selectSingleNode("xMun").Text
        'rsFor!Pais = xEmitEnd.Item(0).selectSingleNode("xPais").Text
        rsFor!Pais = "BRASIL"
        'rsFor!Fone = xEmitEnd.Item(0).selectSingleNode("fone").Text
    
    
    rsFor.Update
    On Error GoTo 0
   
'Apos cadastrar o fornecedor, x buscara o ID desse fornecedor para ser utilizado na importação do xml em questao

'x = Nz(DLookup("idCliente", "tbcliente", "Cnpj = '" & doc.getElementsByTagName("CNPJ_dest")(0).Text & "'"), 0)



X = Nz(DLookup("idFor", "tbfornecedor", "Cnpj = '" & varCNPJ & "'"), 0)     'X buscará o fornecedor na tabela "tbFornecedores"

Else
'atualiza dados do fornecedor já cadastrado
rsFor.Close
Set rsFor = Db.OpenRecordset("SELECT * FROM tbFornecedor WHERE IdFor = " & X & "")
rsFor.Edit
On Error Resume Next
        
        rsFor!Tipo = "TRANSPORTADORA"
        rsFor!CNPJ = xEmit.Item(0).selectSingleNode("CNPJ").Text
        varCNPJ = xEmit.Item(0).selectSingleNode("CNPJ").Text
        
        rsFor!RazaoSocial = xEmit.Item(0).selectSingleNode("xNome").Text
        rsFor!IE = xEmit.Item(0).selectSingleNode("IE").Text
       
        rsFor!CRT = "NAO INFORMADO"
                
        rsFor!Logradouro = xEmitEnd.Item(0).selectSingleNode("xLgr").Text
        rsFor!Nro = xEmitEnd.Item(0).selectSingleNode("nro").Text
        rsFor!CEP = xEmitEnd.Item(0).selectSingleNode("CEP").Text
        
        
        'rsFor!Compl = xDet.Item(0).selectSingleNode("xCpl").Text
        rsFor!Bairro = xEmitEnd.Item(0).selectSingleNode("xBairro").Text
        rsFor!UF = xEmitEnd.Item(0).selectSingleNode("UF").Text
        rsFor!Municipio = xEmitEnd.Item(0).selectSingleNode("xMun").Text
        'rsFor!Pais = xEmitEnd.Item(0).selectSingleNode("xPais").Text
        rsFor!Pais = "BRASIL"
        'rsFor!Fone = xEmitEnd.Item(0).selectSingleNode("fone").Text

rsFor.Update
On Error GoTo 0

End If
rsFor.Close
Set rsFor = Nothing

'------------------------------------------------------------------------'
'Dados Principais da Nota de Venda (tbVendas)
'Set xDet = doc.getElementsByTagName("det")
Set rsTransp = Db.OpenRecordset("tbTransportes")

'verifica se o xml já foi processado antes pra não duplicar a linha
x1 = Nz(DLookup("ChaveCTe", "tbTransportes", "ChaveCTe = '" & doc.getElementsByTagName("chCTe")(0).Text & "'"), 0)
If x1 = doc.getElementsByTagName("chCTe")(0).Text Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar o fornecedor
GoTo proximoarquivo
Else
End If


rsTransp.AddNew
    rsTransp!ID_Emit = X
    'Necessario essa verificação pois na versao XML 1.10 era somente Data (dEmi) ja na 3.0 mudou para DataHora (dhEmi)
    
    If (doc.getElementsByTagName("dhEmi").length) Then
    rsTransp!DataEmissao = Format(Left(doc.getElementsByTagName("dhEmi")(0).Text, 10), "dd/mm/yyyy")
    Else
    rsVenda!DataEmissao = Format(doc.getElementsByTagName("dEmi")(0).Text, "dd/mm/yyyy")
    End If
        'Passa os totais da NFe para a variavel xProd
        'xProd = doc.getElementsByTagName("total")(0).XML
        'Valor Bruto sem desconto
       
       
    rsTransp!Num_CTE = xide.Item(0).selectSingleNode("nCT").Text
    
    rsTransp!Serie = xide.Item(0).selectSingleNode("serie").Text
    rsTransp!ChaveCTE = doc.getElementsByTagName("chCTe")(0).Text
    rsTransp!ValorTotalServico = Replace((doc.getElementsByTagName("vTPrest")(0).Text), ".", ",")
    
    Select Case doc.getElementsByTagName("CST")(0).Text
    Case Is = "00" 'Tributada integralmente...
    rsTransp!CST_Desc = "00 - tributação normal do ICMS"
    Case Is = "40" 'Isento
    rsTransp!CST_Desc = "40 - isento de ICMS"
    Case Is = "90" 'outros'
    rsTransp!CST_Desc = "90 - Outros"
    End Select
    rsTransp!CST = doc.getElementsByTagName("CST")(0).Text


    Set xDestEnd = doc.getElementsByTagName("enderDest")
    teste = xDestEnd.Item(0).selectSingleNode("xLgr").Text
    
    
    'Set xICMS = doc.getElementsByTagName("imp")
    'TESTE2 = xICMS.Item(0).selectSingleNode("vBC").Text
    Dim xICMS As String
    xICMS = doc.getElementsByTagName("imp")(0).XML
    
    rsTransp!BaseCalcICMS = Replace(separaEntreDuasStringsXML(xICMS, "<vBC>", "</vBC>"), ".", ",")
    rsTransp!AliqICMS = Replace(separaEntreDuasStringsXML(xICMS, "<pICMS>", "</pICMS>"), ".", ",")
    rsTransp!ValorICMS = Replace(separaEntreDuasStringsXML(xICMS, "<vICMS>", "</vICMS>"), ".", ",")
    'rsTransp!TotalTributos = Replace(ximp.Item(0).selectSingleNode("vTotTrib").Text, ".", ",")
    rsTransp!CFOP = doc.getElementsByTagName("CFOP")(0).Text
    
    
    Select Case doc.getElementsByTagName("toma")(0).Text
    Case Is = 0
    rsTransp!Tomador = "REMETENTE"
    Case Is = 1
    rsTransp!Tomador = "EXPEDIDOR"
    Case Is = 2
    rsTransp!Tomador = "RECEBEDOR"
    Case Is = 3
    rsTransp!Tomador = "DESTINATARIO"
    Case Is = 4
    rsTransp!Tomador = "OUTROS"
    End Select
    
   
    rsTransp!RemetenteCNPJ = xRem.Item(0).selectSingleNode("CNPJ").Text
    rsTransp!RemetenteRazaoSocial = xRem.Item(0).selectSingleNode("xNome").Text
    rsTransp!RemetenteIE = xRem.Item(0).selectSingleNode("IE").Text
    rsTransp!RemetenteUF = xRemEnd.Item(0).selectSingleNode("UF").Text
    rsTransp!RemetenteCidade = xRemEnd.Item(0).selectSingleNode("xMun").Text
    rsTransp!RemetenteEnd = xRemEnd.Item(0).selectSingleNode("xLgr").Text
    On Error Resume Next
    rsTransp!DestinatarioCNPJ = xDest.Item(0).selectSingleNode("CNPJ").Text
    rsTransp!DestinatarioRazaoSocial = xDest.Item(0).selectSingleNode("xNome").Text
    rsTransp!DestinatarioIE = xDest.Item(0).selectSingleNode("IE").Text
    rsTransp!DestinatarioUF = xDestEnd.Item(0).selectSingleNode("UF").Text
    rsTransp!DestinatarioCidade = xDestEnd.Item(0).selectSingleNode("xMun").Text
    rsTransp!DestinatarioEnd = xDestEnd.Item(0).selectSingleNode("xLgr").Text
       
    
    
rsTransp.Update
On Error GoTo 0
regAtual = Nz(DLookup("ID", "tbTransportes", "ChaveCTe = '" & doc.getElementsByTagName("chCTe")(0).Text & "'"), 0)



rsTransp.Close
Set rsTransp = Nothing

 'LANÇAR NO CONTAS A PAGAR
 'REMETENTE - PAGADOR
    strSQL = ("INSERT INTO tb_Detalhe_Boletos_Compras ( DtEmissao, ValorOriginal, NumBoleto, Id_Fornecedor, Fornecedor, Chave_Nfe, STATUS ) " & _
                    "SELECT tbTransportes.DataEmissao, tbTransportes.ValorTotalServico, 'LANC AUTOMATICO' AS Num_boleto, tbTransportes.ID_Emit, tbFornecedor.RazaoSocial, tbTransportes.ChaveCTE, 'ABERTO' AS STATUS " & _
                    "FROM tbFornecedor INNER JOIN tbTransportes ON tbFornecedor.IDFor = tbTransportes.ID_Emit " & _
                    "WHERE (((tbTransportes.Tomador)='REMETENTE') AND ((tbTransportes.RemetenteCNPJ)='23866944000141') AND ((tbTransportes.ChaveCTE)='" & xide.Item(0).selectSingleNode("nCT").Text & "'));")
    Conn.Execute strSQL
  'DESTINATARIO - PAGADOR
  strSQL = ("INSERT INTO tb_Detalhe_Boletos_Compras ( DtEmissao, ValorOriginal, NumBoleto, Id_Fornecedor, Fornecedor, Chave_Nfe, STATUS ) " & _
                "SELECT tbTransportes.DataEmissao, tbTransportes.ValorTotalServico, 'LANC AUTOMATICO' AS Num_boleto, tbTransportes.ID_Emit, tbFornecedor.RazaoSocial, tbTransportes.ChaveCTE, 'ABERTO' AS STATUS " & _
                "FROM tbFornecedor INNER JOIN tbTransportes ON tbFornecedor.IDFor = tbTransportes.ID_Emit " & _
                "WHERE (((tbTransportes.Tomador)='DESTINATARIO') AND ((tbTransportes.DestinatarioCNPJ)='23866944000141') AND ((tbTransportes.ChaveCTE)='" & xide.Item(0).selectSingleNode("nCT").Text & "'));")
  Conn.Execute strSQL
    'LANCAR NO CONTAS A PAGAR
'------------------------------------------------------------------------'
' Dados dos Produtos - chaves de nfes transportadas

'verifica se o xml já foi processado antes pra não duplicar a linha
X = Nz(DLookup("ID", "tbTransportesDet", "ID = " & regAtual & ""), 0)
If X = regAtual Then 'Se x for <=0 significa que nao ta cadastrado, entao irá cadastrar a chave de transportes
GoTo proximoarquivo
Else
End If


I = 0

'Aqui é o Loop que percorrerá pela Tag "det" que são os produtos..
'Buscara produto a produto, e o inserirá na nota que esta sendo importada
For Each CHAVE In xDet

xinfDoc = doc.getElementsByTagName("infNFe")(I).XML ' xProd desmembrará o xml pegando produto a produto...
X = Nz(DLookup("ID", "tbTransportesDet", "ChaveNFe = '" & separaEntreDuasStringsXML(Replace(xinfDoc, "'", ""), "<chave>", "</chave>") & "'"), 0)
If X <= 0 Then
'Cadastra o Produto, pois ainda nao foi cadastrado
Set rsTrDet = Db.OpenRecordset("tbTransportesDet")
rsTrDet.AddNew
    rsTrDet!ID_Transporte = regAtual
    rsTrDet!chaveNFE = separaEntreDuasStringsXML(Replace(xinfDoc, "'", ""), "<chave>", "</chave>")
    
   rsTrDet.Update
   
'Insere o produto cadastrado na Nota de compra


rsTrDet.Close

Set rsTrDet = Nothing

Else
rsTrDet.Close
Set rsTrDet = Nothing
    
     
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

strSQL = ("UPDATE tbTransportes SET tbTransportes.CST = Left(tbTransportes.CST_Desc,2) where tbTransportes.CST is null;")
Conn.Execute strSQL

Forms!frmXMLinput.Requery
MsgBox "Todos os XMLs da pasta selecionada Foram Importados com Sucesso! " & vbNewLine & "Verifique as Notas lançadas!!!", vbInformation, "Sucesso!!!"
Db.Close
Set Db = Nothing
Call DisconnectFromDataBase
End Function



