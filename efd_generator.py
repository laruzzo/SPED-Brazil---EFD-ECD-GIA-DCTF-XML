import datetime
import os
from db_connection import DBConnection

class EFDGenerator:
    def __init__(self):
        self.db_conn = DBConnection()

    def generate_efd_icms_ipi(self, cDtIni, cDtFim, clocal, cDtINI_Contabil, cIDIventario):
        self.db_conn.connect_to_database()

        # Formatação de datas
        dt_ini = datetime.datetime.strptime(cDtIni, "%Y-%m-%d")
        dt_fim = datetime.datetime.strptime(cDtFim, "%Y-%m-%d")

        cSTR_DtINI = dt_ini.strftime("%d%m%Y")
        cSTR_DtFIM = dt_fim.strftime("%d%m%Y")

        cDtIniVb = dt_ini.strftime("%m/%d/%Y")
        cDtFimVb = dt_fim.strftime("%m/%d/%Y")

        # Nome do arquivo de saída
        output_filename = os.path.join(clocal, f"EFD_ICMS_IPI_{dt_ini.month}_{dt_ini.year}.txt")

        clintot = 0 # Contador de linhas totais

        with open(output_filename, 'w', encoding='utf-8') as f:
            # Bloco 0 - Abertura, Identificação e Referências

            # Determinar a versão do código
            cCodVer = "012"
            if dt_ini >= datetime.datetime(2019, 1, 1):
                cCodVer = "015"
            if dt_ini >= datetime.datetime(2022, 1, 1):
                cCodVer = "016"
            if dt_ini >= datetime.datetime(2023, 1, 1):
                cCodVer = "017"

            cCodFin = "0"
            cPerfil = "A"
            cAtividade = "0"

            # 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA PESSOA JURÍDICA
            rsEmpresa = self.db_conn.fetch_all("SELECT * FROM tbEmpresa")
            if rsEmpresa:
                empresa = rsEmpresa[0]
                c0000 = f"|0000|{cCodVer}|{cCodFin}|{cSTR_DtINI}|{cSTR_DtFIM}|{empresa['RazaoSocial']}|{empresa['CNPJ']}|{empresa['UF']}|{empresa['IE']}|{empresa['Cod_Mun']}|{empresa['IM']}|{empresa['Suframa']}|{cPerfil}|{cAtividade}|"
                f.write(c0000 + '\n')
                clintot += 1

            # 0001: ABERTURA DO BLOCO 0
            c0001 = "|0001|0|"
            f.write(c0001 + '\n')
            clintot += 1

            # 0100: DADOS DO CONTABILISTA
            rsContador = self.db_conn.fetch_all("SELECT * FROM tbContador")
            if rsContador:
                contador = rsContador[0]
                c0100 = f"|0100|{contador['Nome']}|{contador['CPF']}|{contador['CRC']}|{contador['CNPJ']}|{contador['CEP']}|{contador['Endereco']}|{contador['Num']}|{contador['Compl']}|{contador['Bairro']}|{contador['Fone']}|{contador['Fax']}|{contador['Email']}|{contador['Cod_Mun']}|"
                f.write(c0100 + '\n')
                clintot += 1

            # 0150: TABELA DE CADASTRO DO PARTICIPANTE
            # Clientes
            rsCliente = self.db_conn.fetch_all(f"SELECT tbCliente.IDCliente, tbCliente.Tipo, tbCliente.Cnpj, tbCliente.RazaoSocial, tbCliente.IE, tbCliente.CRT, tbCliente.CEP, tbCliente.Logradouro, tbCliente.Nro, tbCliente.Compl, tbCliente.Bairro, tbCliente.UF, tbCliente.cod_Municipio, tbCliente.Municipio, tbCliente.Pais, tbCliente.Fone, tbCliente.Email FROM tbCliente INNER JOIN tbVendas ON tbCliente.IDCliente = tbVendas.IdCliente WHERE (((tbVendas.DataEmissao) >= '#{cDtIniVb}#' And (tbVendas.DataEmissao) <= '#{cDtFimVb} 23:59:59#')) GROUP BY tbCliente.IDCliente, tbCliente.Tipo, tbCliente.Cnpj, tbCliente.RazaoSocial, tbCliente.IE, tbCliente.CRT, tbCliente.CEP, tbCliente.Logradouro, tbCliente.Nro, tbCliente.Compl, tbCliente.Bairro, tbCliente.UF, tbCliente.cod_Municipio, tbCliente.Municipio, tbCliente.Pais, tbCliente.Fone, tbCliente.Email;")
            if rsCliente:
                for cliente in rsCliente:
                    c0150 = f"|0150|{cliente['IDCliente']}|{cliente['Tipo']}|{cliente['Cnpj']}|{cliente['IE']}|{cliente['RazaoSocial']}|{cliente['cod_Municipio']}|{cliente['Logradouro']}|{cliente['Nro']}|{cliente['Compl']}|{cliente['Bairro']}|"
                    f.write(c0150 + '\n')
                    clintot += 1

            # Fornecedores
            # Limpa e insere fornecedores ativos (compras e imobilizado)
            self.db_conn.execute_query("delete from tbFornecedor_Ativo_temp")
            self.db_conn.execute_query(f"insert into tbFornecedor_Ativo_temp SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email FROM tbFornecedor INNER JOIN (tbImobilizado INNER JOIN tbCompras ON tbImobilizado.ChaveNFe = tbCompras.ChaveNF) ON tbFornecedor.IDFor = tbCompras.IdFornecedor WHERE (((tbImobilizado.DataEmissao) >= '{cDtIni}' And (tbImobilizado.DataEmissao) <= '{cDtFim}')) GROUP BY tbCompras.IdFornecedor, tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email; ")
            self.db_conn.execute_query(f"insert into tbFornecedor_Ativo_temp SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email FROM tbFornecedor INNER JOIN tbCompras ON tbFornecedor.IDFor = tbCompras.IdFornecedor LEFT OUTER JOIN tbFornecedor_Ativo_temp ON  tbFornecedor.IDFor = tbFornecedor_Ativo_temp.IDFor WHERE tbCompras.DataEmissao >= '{cDtIni}' And tbCompras.DataEmissao <= '{cDtFim}' and tbFornecedor_Ativo_temp.IDFor is null GROUP BY tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email")
            self.db_conn.execute_query(f"INSERT INTO tbFornecedor_Ativo_temp ( IDFor, Tipo, Cnpj, RazaoSocial, IE, CRT, CEP, Logradouro, Nro, Compl, Bairro, UF, cod_Municipio, Municipio, Pais, Fone, Email ) SELECT tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email FROM tbFornecedor INNER JOIN tbTransportes ON tbFornecedor.IDFor = tbTransportes.ID_Emit where tbTransportes.LancFiscal = 'CREDITO' GROUP BY tbFornecedor.IDFor, tbFornecedor.Tipo, tbFornecedor.Cnpj, tbFornecedor.RazaoSocial, tbFornecedor.IE, tbFornecedor.CRT, tbFornecedor.CEP, tbFornecedor.Logradouro, tbFornecedor.Nro, tbFornecedor.Compl, tbFornecedor.Bairro, tbFornecedor.UF, tbFornecedor.cod_Municipio, tbFornecedor.Municipio, tbFornecedor.Pais, tbFornecedor.Fone, tbFornecedor.Email, tbTransportes.DataEmissao HAVING tbTransportes.DataEmissao>='{cDtIni}' And tbTransportes.DataEmissao<='{cDtFim}';")
            self.db_conn.execute_query("DELETE FROM tbFornecedor_Ativo_temp WHERE tbFornecedor_Ativo_temp.IDFor=1136")

            rsFornecedor = self.db_conn.fetch_all("SELECT tbFornecedor_Ativo_temp.IDFor, tbFornecedor_Ativo_temp.Tipo, tbFornecedor_Ativo_temp.Cnpj, tbFornecedor_Ativo_temp.RazaoSocial, tbFornecedor_Ativo_temp.IE, tbFornecedor_Ativo_temp.CRT, tbFornecedor_Ativo_temp.CEP, tbFornecedor_Ativo_temp.Logradouro, tbFornecedor_Ativo_temp.Nro, tbFornecedor_Ativo_temp.Compl, tbFornecedor_Ativo_temp.Bairro, tbFornecedor_Ativo_temp.UF, tbFornecedor_Ativo_temp.cod_Municipio, tbFornecedor_Ativo_temp.Municipio, tbFornecedor_Ativo_temp.Pais, tbFornecedor_Ativo_temp.Fone, tbFornecedor_Ativo_temp.Email FROM tbFornecedor_Ativo_temp GROUP BY tbFornecedor_Ativo_temp.IDFor, tbFornecedor_Ativo_temp.Tipo, tbFornecedor_Ativo_temp.Cnpj, tbFornecedor_Ativo_temp.RazaoSocial, tbFornecedor_Ativo_temp.IE, tbFornecedor_Ativo_temp.CRT, tbFornecedor_Ativo_temp.CEP, tbFornecedor_Ativo_temp.Logradouro, tbFornecedor_Ativo_temp.Nro, tbFornecedor_Ativo_temp.Compl, tbFornecedor_Ativo_temp.Bairro, tbFornecedor_Ativo_temp.UF, tbFornecedor_Ativo_temp.cod_Municipio, tbFornecedor_Ativo_temp.Municipio, tbFornecedor_Ativo_temp.Pais, tbFornecedor_Ativo_temp.Fone, tbFornecedor_Ativo_temp.Email;")
            if rsFornecedor:
                for fornecedor in rsFornecedor:
                    c0150 = f"|0150|{fornecedor['IDFor']}|{fornecedor['Tipo']}|{fornecedor['Cnpj']}|{fornecedor['IE']}|{fornecedor['RazaoSocial']}|{fornecedor['cod_Municipio']}|{fornecedor['Logradouro']}|{fornecedor['Nro']}|{fornecedor['Compl']}|{fornecedor['Bairro']}|"
                    f.write(c0150 + '\n')
                    clintot += 1

            # 0200: TABELA DE IDENTIFICAÇÃO DO ITEM (PRODUTO E SERVIÇOS)
            self.db_conn.execute_query("delete from tbCadProd_Ativo_temp")
            self.db_conn.execute_query(f"INSERT INTO tbCadProd_Ativo_temp ( IDProd ) SELECT tbCadProd.IDProd FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON tbCadProd.IDProd = tbVendasDet.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda WHERE (((tbVendas.DataEmissao) >= '{cDtIni}' And (tbVendas.DataEmissao) <= '{cDtFim} 23:59:59')) GROUP BY tbCadProd.IDProd;")
            self.db_conn.execute_query(f"INSERT INTO tbCadProd_Ativo_temp ( IDProd ) SELECT tbCadProd.IDProd FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra WHERE (((tbCompras.DataEmissao) >= '{cDtIni}' And (tbCompras.DataEmissao) <= ' {cDtFim} ')) GROUP BY tbCadProd.IDProd;")
            self.db_conn.execute_query(f"INSERT INTO tbCadProd_Ativo_temp ( IDProd ) SELECT tbCadProd.IDProd FROM tbVendas, tbCadProd INNER JOIN tbImobilizado ON tbCadProd.IDProd = tbImobilizado.IDProd LEFT OUTER JOIN  tbCadProd_Ativo_temp ON tbCadProd.IDProd = tbCadProd_Ativo_temp.IDProd WHERE tbImobilizado.DataEmissao >= '{cDtIni}' And tbImobilizado.DataEmissao <= '{cDtFim}' and  tbCadProd_Ativo_temp.IDProd is null GROUP BY tbCadProd.IDProd;")
            self.db_conn.execute_query(f"INSERT INTO tbCadProd_Ativo_temp ( IDProd ) SELECT tbIventarioDet.ID_Prod FROM tbIventarioDet LEFT JOIN tbCadProd_Ativo_temp ON tbIventarioDet.ID_Prod = tbCadProd_Ativo_temp.IDProd GROUP BY tbCadProd_Ativo_temp.IDProd, tbIventarioDet.ID_Prod, tbIventarioDet.ID_Iventario HAVING (((tbCadProd_Ativo_temp.IDProd) Is Null) AND ((tbIventarioDet.ID_Iventario)={cIDIventario}));")
            self.db_conn.execute_query(f"INSERT INTO tbCadProd_Ativo_temp ( IDProd ) SELECT tb_Registro_Consumo.ID_Produto FROM (tb_Registro_Envase INNER JOIN tb_Registro_Consumo ON tb_Registro_Envase.ID = tb_Registro_Consumo.ID_Lote) LEFT JOIN tbCadProd_Ativo_temp ON tb_Registro_Consumo.ID_Produto = tbCadProd_Ativo_temp.IDProd WHERE (((tb_Registro_Envase.DATA)>='{cDtIni}' And (tb_Registro_Envase.DATA)<='{cDtFim}') AND ((tbCadProd_Ativo_temp.IDProd) Is Null));")
            self.db_conn.execute_query(f"INSERT INTO tbCadProd_Ativo_temp ( IDProd ) SELECT tb_Registro_Envase.ID_Produto FROM tb_Registro_Envase LEFT JOIN tbCadProd_Ativo_temp ON tb_Registro_Envase.ID_Produto = tbCadProd_Ativo_temp.IDProd WHERE (((tb_Registro_Envase.DATA)>='{cDtIni}' And (tb_Registro_Envase.DATA)<='{cDtFim}') AND ((tbCadProd_Ativo_temp.IDProd) Is Null));")
            self.db_conn.execute_query("INSERT INTO tbCadProd_Ativo_temp (IDPROD) select CodBem from (SELECT tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.Status, tbCadProd_Ativo_temp.IDProd FROM tbCadProd_Ativo_temp right outer JOIN tbImobilizadoCadastro ON tbCadProd_Ativo_temp.IDProd = tbImobilizadoCadastro.CodBem GROUP BY tbImobilizadoCadastro.CodBem HAVING tbCadProd_Ativo_temp.IDProd Is  Null and tbImobilizadoCadastro.CodBem is not null ) as q1;")
            self.db_conn.execute_query("delete tbCadProd_Ativo_temp from tbCadProd_Ativo_temp inner join (select * from tbimobilizadocadastro where status = 'EXAURIDO') AS Q2 ON tbCadProd_Ativo_temp.IDProd = q2.IDProd;")

            rsCadProd = self.db_conn.fetch_all("SELECT tbCadProd.* FROM tbCadProd_Ativo_temp INNER JOIN tbCadProd ON tbCadProd_Ativo_temp.IDProd = tbCadProd.IDProd;")
            if rsCadProd:
                for produto in rsCadProd:
                    c0200 = f"|0200|{produto['IDProd']}|{produto['Descricao']}|{produto['Unid']}|{produto['TipoItem']}|{produto['NCM']}|{produto['EX_IPI']}|{produto['CodBarra']}|{produto['CodAntItem']}|"
                    f.write(c0200 + '\n')
                    clintot += 1

            # 0300: CADASTRO DE BENS OU COMPONENTES DO ATIVO IMOBILIZADO
            rsImob = self.db_conn.fetch_all(f"SELECT tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza FROM tbPlanoContasContabeis INNER JOIN ((tbImobilizado INNER JOIN tbCadProd_Ativo_temp ON tbImobilizado.IDProd = tbCadProd_Ativo_temp.IDProd) INNER JOIN tbImobilizadoCadastro ON tbImobilizado.IDProd = tbImobilizadoCadastro.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta GROUP BY tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza UNION select distinct q2.* from (SELECT tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza FROM tbPlanoContasContabeis INNER JOIN ((tbImobilizado INNER JOIN tbCadProd_Ativo_temp ON tbImobilizado.IDProd = tbCadProd_Ativo_temp.IDProd) INNER JOIN tbImobilizadoCadastro ON tbImobilizado.IDProd = tbImobilizadoCadastro.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta GROUP BY tbImobilizado.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza HAVING (((tbImobilizadoCadastro.Bem_Componente) = 'COMP')) ) as q1 INNER Join (SELECT tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza FROM tbPlanoContasContabeis INNER JOIN (tbImobilizadoCadastro INNER JOIN tbCadProd_Ativo_temp ON tbImobilizadoCadastro.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbPlanoContasContabeis.ID = tbImobilizadoCadastro.ID_Conta GROUP BY tbImobilizadoCadastro.IDProd, tbImobilizadoCadastro.Bem_Componente, tbImobilizadoCadastro.Descricao, tbImobilizadoCadastro.CodBem, tbImobilizadoCadastro.ID_Conta, tbImobilizadoCadastro.Nr_Parcelas, tbImobilizadoCadastro.Centro_Custo, tbPlanoContasContabeis.Cod_Natureza HAVING(tbImobilizadoCadastro.Bem_Componente) = 'BEM' ) AS Q2 ON Q1.CodBem = Q2.IdProd")
            if rsImob:
                for imobilizado in rsImob:
                    cBemComp = "1" if imobilizado['Bem_Componente'] == "BEM" else "2"
                    c0300 = f"|0300|{imobilizado['IDProd']}|{cBemComp}|{imobilizado['Descricao']}|{imobilizado['CodBem']}|{imobilizado['ID_Conta']}|{imobilizado['Nr_Parcelas']}|"
                    f.write(c0300 + '\n')
                    clintot += 1

                    cBemCC = "3" # Default para Area Produtiva
                    cBemCCDesc = "Area Produtiva"
                    if imobilizado['Centro_Custo'] == "ADM":
                        cBemCC = "5"
                        cBemCCDesc = "Area Administrativa"
                    c0305 = f"|0305|{cBemCC}|{cBemCCDesc}|{imobilizado['Nr_Parcelas']}|"
                    f.write(c0305 + '\n')
                    clintot += 1

            # 0400: TABELA DE NATUREZA DA OPERAÇÃO/PRESTAÇÃO
            rsCFOP = self.db_conn.fetch_all(f"SELECT tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC FROM tbCompras INNER JOIN (tbComprasDet INNER JOIN tbCadProd_Ativo_temp ON tbComprasDet.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbCompras.ID = tbComprasDet.IDCompra WHERE (((tbCompras.IdFornecedor) <> 1131)) GROUP BY tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.CFOP_ESC_DESC UNION SELECT tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC FROM tbVendas INNER JOIN (tbVendasDet INNER JOIN tbCadProd_Ativo_temp ON tbVendasDet.IDProd = tbCadProd_Ativo_temp.IDProd) ON tbVendas.ID = tbVendasDet.IDVenda GROUP BY tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.CFOP_ESC_DESC, tbVendas.TipoNF HAVING (((tbVendasDet.CFOP_ESCRITURADA) Is Not Null) AND ((tbVendas.TipoNF)='0-ENTRADA'));")
            if rsCFOP:
                for cfop_item in rsCFOP:
                    if cfop_item['CFOP_ESCRITURADA'] != 1252:
                        c0400 = f"|0400|{cfop_item['CFOP_ESCRITURADA']}|{cfop_item['CFOP_ESC_DESC']}|"
                        f.write(c0400 + '\n')
                        clintot += 1

            # 0500: PLANO DE CONTAS CONTÁBEIS
            rsContas = self.db_conn.fetch_all("SELECT * FROM tbPlanoContasContabeis")
            if rsContas:
                for conta in rsContas:
                    c0500 = f"|0500|{cSTR_DtINI}|{conta['Cod_Natureza']}|{conta['Cod_Indicador']}|1|{conta['ID']}|{conta['Desc_CodNatureza']}|"
                    f.write(c0500 + '\n')
                    clintot += 1

            # 0600: CENTRO DE CUSTOS
            c0600_prod = f"|0600|{cSTR_DtINI}|3|área produtiva|"
            f.write(c0600_prod + '\n')
            clintot += 1
            c0600_adm = f"|0600|{cSTR_DtINI}|5|área administrativa|"
            f.write(c0600_adm + '\n')
            clintot += 1

            # Bloco B - ISS
            cB001 = "|B001|1|"
            f.write(cB001 + '\n')
            clintot += 1

            # Bloco C - Documentos Fiscais I (ICMS/IPI)
            cC001 = "|C001|0|" # Abertura do Bloco C
            f.write(cC001 + '\n')
            clintot += 1

            # C100: NOTA FISCAL (código 01, 1B, 04)
            rsCompra = self.db_conn.fetch_all(f"SELECT tbCompras.ID, tbCompras.IdFornecedor, tbCompras.Serie, tbCompras.NumNF, tbCompras.ChaveNF, tbCompras.DataEmissao, tbCompras.VlrTOTALNF, tbCompras.VlrDesconto, tbCompras.VlrTotalProdutos, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.BaseCalculo,0)) AS BaseCalculo, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_ICMS,0)) AS Valor_ICMS, Sum(tbComprasDet.BaseCalc_ST) AS BaseCalc_ST, Sum(tbComprasDet.Valor_ICMS_ST) AS Valor_ICMS_ST, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_IPI,0)) AS Valor_IPI, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_PIS,0)) AS Valor_PIS, Sum(IIf(lancfiscal='CREDITO',tbComprasDet.Valor_Cofins,0)) AS Valor_Cofins FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra GROUP BY tbCompras.ID, tbCompras.IdFornecedor, tbCompras.Serie, tbCompras.NumNF, tbCompras.ChaveNF, tbCompras.DataEmissao, tbCompras.VlrTOTALNF, tbCompras.VlrDesconto, tbCompras.VlrTotalProdutos, tbCompras.dataemissao HAVING tbCompras.dataemissao>= '#{cDtIni}#' and dataemissao <= '#{cDtFim}#' and IdFornecedor <> 1131;")
            if rsCompra:
                for compra in rsCompra:
                    c100 = f"|C100|0|1|{compra['IdFornecedor']}|01|00|{compra['Serie']}|{compra['NumNF']}|{compra['ChaveNF']}|{compra['DataEmissao'].strftime('%d%m%Y')}|{compra['DataEmissao'].strftime('%d%m%Y')}|{compra['VlrTOTALNF']}|{compra['VlrDesconto']}|{compra['VlrTotalProdutos']}|{compra['BaseCalculo']}|{compra['Valor_ICMS']}|{compra['BaseCalc_ST']}|{compra['Valor_ICMS_ST']}|{compra['Valor_IPI']}|{compra['Valor_PIS']}|{compra['Valor_Cofins']}|"
                    f.write(c100 + '\n')
                    clintot += 1

                    # C170: ITENS DO DOCUMENTO
                    rsCompraDet = self.db_conn.fetch_all(f"SELECT tbComprasDet.ID as ID_DET, tbComprasDet.IDProd, tbComprasDet.Qnt, tbCadProd.Unid, tbComprasDet.ValorTot, tbComprasDet.VlrDesc, tbComprasDet.CST, tbComprasDet.CFOP, tbComprasDet.CFOP_ESCRITURADA, tbComprasDet.BaseCalculo, tbComprasDet.Aliq_ICMS, tbComprasDet.Valor_ICMS, tbComprasDet.BaseCalc_ST, tbComprasDet.Aliq_ICMS_ST, tbComprasDet.Valor_ICMS_ST, tbComprasDet.Aliq_IPI, tbComprasDet.Valor_IPI, tbComprasDet.Aliq_PIS, tbComprasDet.Aliq_Cofins, tbComprasDet.Valor_PIS, tbComprasDet.Valor_Cofins FROM tbCompras INNER JOIN (tbCadProd INNER JOIN tbComprasDet ON (tbCadProd.IDProd = tbComprasDet.IDProd) AND (tbCadProd.IDProd = tbComprasDet.IDProd)) ON tbCompras.ID = tbComprasDet.IDCompra WHERE tbCompras.ID = {compra['ID']} ORDER BY tbComprasDet.ID;")
                    if rsCompraDet:
                        for item_compra in rsCompraDet:
                            c170 = f"|C170|{item_compra['ID_DET']}|{item_compra['IDProd']}|{item_compra['Qnt']}|{item_compra['Unid']}|{item_compra['ValorTot']}|{item_compra['VlrDesc']}|{item_compra['CST']}|{item_compra['CFOP']}|{item_compra['CFOP_ESCRITURADA']}|{item_compra['BaseCalculo']}|{item_compra['Aliq_ICMS']}|{item_compra['Valor_ICMS']}|{item_compra['BaseCalc_ST']}|{item_compra['Aliq_ICMS_ST']}|{item_compra['Valor_ICMS_ST']}|{item_compra['Aliq_IPI']}|{item_compra['Valor_IPI']}|{item_compra['Aliq_PIS']}|{item_compra['Aliq_Cofins']}|{item_compra['Valor_PIS']}|{item_compra['Valor_Cofins']}|"
                            f.write(c170 + '\n')
                            clintot += 1

            rsVenda = self.db_conn.fetch_all(f"select * from tbVendas where dataemissao >= '#{cDtIni}#' and dataemissao <= '#{cDtFim}#' and tipoNF = '1-SAIDA' and [Status] = 'ATIVO' and NatOperacao <> 'Venda Cupom Fiscal SAT'")
            if rsVenda:
                for venda in rsVenda:
                    c100 = f"|C100|1|1|{venda['IdCliente']}|01|00|{venda['Serie']}|{venda['NumNF']}|{venda['ChaveNF']}|{venda['DataEmissao'].strftime('%d%m%Y')}|{venda['DataEmissao'].strftime('%d%m%Y')}|{venda['VlrTOTALNF']}|{venda['VlrDesconto']}|{venda['VlrTotalProdutos']}|{venda['BaseCalculo']}|{venda['Valor_ICMS']}|{venda['BaseCalc_ST']}|{venda['Valor_ICMS_ST']}|{venda['Valor_IPI']}|{venda['Valor_PIS']}|{venda['Valor_Cofins']}|"
                    f.write(c100 + '\n')
                    clintot += 1

                    # C170: ITENS DO DOCUMENTO
                    rsVendaDet = self.db_conn.fetch_all(f"SELECT tbVendasDet.ID as ID_DET, tbVendasDet.IDProd, tbVendasDet.Qnt, tbCadProd.Unid, tbVendasDet.ValorTot, tbVendasDet.VlrDesc, tbVendasDet.CST, tbVendasDet.CFOP, tbVendasDet.CFOP_ESCRITURADA, tbVendasDet.BaseCalculo, tbVendasDet.Aliq_ICMS, tbVendasDet.Valor_ICMS, tbVendasDet.BaseCalc_ST, tbVendasDet.Aliq_ICMS_ST, tbVendasDet.Valor_ICMS_ST, tbVendasDet.Aliq_IPI, tbVendasDet.Valor_IPI, tbVendasDet.Aliq_PIS, tbVendasDet.Aliq_Cofins, tbVendasDet.Valor_PIS, tbVendasDet.Valor_Cofins FROM tbVendas INNER JOIN (tbCadProd INNER JOIN tbVendasDet ON (tbCadProd.IDProd = tbVendasDet.IDProd) AND (tbCadProd.IDProd = tbVendasDet.IDProd)) ON tbVendas.ID = tbVendasDet.IDVenda WHERE tbVendas.ID = {venda['ID']} ORDER BY tbVendasDet.ID;")
                    if rsVendaDet:
                        for item_venda in rsVendaDet:
                            c170 = f"|C170|{item_venda['ID_DET']}|{item_venda['IDProd']}|{item_venda['Qnt']}|{item_venda['Unid']}|{item_venda['ValorTot']}|{item_venda['VlrDesc']}|{item_venda['CST']}|{item_venda['CFOP']}|{item_venda['CFOP_ESCRITURADA']}|{item_venda['BaseCalculo']}|{item_venda['Aliq_ICMS']}|{item_venda['Valor_ICMS']}|{item_venda['BaseCalc_ST']}|{item_venda['Aliq_ICMS_ST']}|{item_venda['Valor_ICMS_ST']}|{item_venda['Aliq_IPI']}|{item_venda['Valor_IPI']}|{item_venda['Aliq_PIS']}|{item_venda['Aliq_Cofins']}|{item_venda['Valor_PIS']}|{item_venda['Valor_Cofins']}|"
                            f.write(c170 + '\n')
                            clintot += 1

            # C190: REGISTRO ANALÍTICO DO DOCUMENTO
            rsRegSaida = self.db_conn.fetch_all(f"SELECT tbvendas.NatOperacao, Year(DataEmissao) AS ANO, Month(DataEmissao) AS MES, tbvendasdet.CFOP_ESCRITURADA AS CFOP, tbvendasdet.CFOP_ESC_DESC AS [CFOP Desc], tbvendasdet.lancfiscal AS [Lanc Fiscal], Sum(tbvendasdet.ValorTot) AS [Valor Contabil], Sum(tbvendasdet.BaseCalculo) AS [Base de Calculo], Sum(tbvendasdet.Valor_ICMS) AS ICMS, Sum(tbvendasdet.Valor_IPI) AS IPI, Sum(tbvendasdet.Valor_PIS) AS PIS, Sum(tbvendasdet.Valor_Cofins) AS Cofins, Sum(tbvendasdet.Valor_ICMS_ST) AS [ICMS ST], Sum(tbvendasdet.BaseCalc_ST) AS BaseCalc_ST, tbvendasdet.CST, tbvendasdet.CST_DESC, tbvendasdet.Aliq_ICMS FROM (tbCliente INNER JOIN tbvendas ON tbCliente.IDCliente = tbvendas.Idcliente) INNER JOIN (tbCadProd INNER JOIN tbvendasdet ON (tbCadProd.IDProd = tbvendasdet.IDProd) AND (tbCadProd.IDProd = tbvendasdet.IDProd)) ON tbvendas.ID = tbvendasdet.IDVenda WHERE (((tbVendas.DataEmissao) >= '#{cDtIni}#' And (tbVendas.DataEmissao) <= '#{cDtFim}#')) GROUP BY tbvendas.NatOperacao, Year(DataEmissao), Month(DataEmissao), tbvendasdet.CFOP_ESCRITURADA, tbvendasdet.CFOP_ESC_DESC, tbvendasdet.lancfiscal, tbvendas.TipoNF, tbvendasdet.CST, tbvendasdet.CST_DESC, tbvendasdet.Aliq_ICMS HAVING tbvendas.TipoNF = '1-SAIDA' and tbVendas.NatOperacao <> 'Venda Cupom Fiscal SAT' ORDER BY Year(DataEmissao), Month(DataEmissao);")
            if rsRegSaida:
                for reg_saida in rsRegSaida:
                    c190 = f"|C190|{reg_saida['CST']}|{reg_saida['CFOP']}|{reg_saida['Aliq_ICMS']}|{round(reg_saida['Valor Contabil'], 2)}|{round(reg_saida['Base de Calculo'], 2)}|{round(reg_saida['ICMS'], 2)}|{round(reg_saida['BaseCalc_ST'], 2)}|{round(reg_saida['ICMS ST'], 2)}|0|{round(reg_saida['IPI'], 2)}||"
                    f.write(c190 + '\n')
                    clintot += 1

            # C500: NOTA FISCAL/CONTA DE ENERGIA ELÉTRICA (CÓDIGO 06)
            rsEnergia = self.db_conn.fetch_all(f"SELECT * FROM tbCompras WHERE (((tbCompras.IdFornecedor)=1131) AND ((tbCompras.DataEmissao)>= '#{cDtIni}#' And (tbCompras.DataEmissao)<= '#{cDtFim}#'));")
            if rsEnergia:
                for energia in rsEnergia:
                    c500 = f"|C500|0|1|{energia['IdFornecedor']}|06|00|{energia['Serie']}|{energia['NumNF']}|{energia['DataEmissao'].strftime('%d%m%Y')}|{energia['DataEmissao'].strftime('%d%m%Y')}|{energia['VlrTOTALNF']}|{energia['VlrDesconto']}|{energia['VlrTotalProdutos']}|0|{energia['ICMS_BaseCalc']}|{energia['ICMS_Valor']}|0|0||{energia['PIS_Valor']}|{energia['COFINS_Valor']}|2|12|||||||||"
                    f.write(c500 + '\n')
                    clintot += 1

                    # C590: CONSOLIDAÇÃO DE NF DE ENERGIA
                    rsEnergiaDet = self.db_conn.fetch_all(f"SELECT tbComprasDet.* FROM tbCompras INNER JOIN tbComprasDet ON tbCompras.ID = tbComprasDet.IDCompra WHERE tbCompras.ID = {energia['ID']};")
                    if rsEnergiaDet:
                        for energia_det in rsEnergiaDet:
                            c590 = f"|C590|{energia_det['CST_ICMS']}|{energia_det['CFOP_ESCRITURADA']}|{energia_det['Aliq_ICMS']}|{energia_det['ValorTot']}|{energia_det['BaseCalculo']}|{energia_det['Valor_ICMS']}|0|0|0||"
                            f.write(c590 + '\n')
                            clintot += 1

            # C800: CUPOM FISCAL ELETRÔNICO – SAT (CF-E-SAT) (CÓDIGO 59)
            rsSAT = self.db_conn.fetch_all(f"select * from tbVendasSAT WHERE tbVendasSAT.DataEmissao  >= '#{cDtIni} 00:00:00#' And tbVendasSAT.DataEmissao <= '#{cDtFim} 23:59:59#';")
            if rsSAT:
                for sat in rsSAT:
                    c800 = f"|C800|59|00|{sat['numCF']}|{sat['DataEmissao'].strftime('%d%m%Y')}|{sat['Vlr_TotalCF']}|{sat['vPIS']}|{sat['vCofins']}|{sat['CPF_CNPJ']}|{sat['NumSerieSAT']}|{sat['ChaveCF']}|{sat['vDesc']}|{sat['Vlr_TotalCF']}|0|{sat['vICMS']}|0|0|"
                    f.write(c800 + '\n')
                    clintot += 1

                    # C810: ITENS DO CUPOM FISCAL ELETRÔNICO – SAT
                    rsSATDet = self.db_conn.fetch_all(f"select * from tbVendasSATDet where idSAT = {sat['idSAT']}")
                    if rsSATDet:
                        for sat_det in rsSATDet:
                            c810 = f"|C810|{sat_det['NumItem']}|{sat_det['IDProd']}|{sat_det['Qt']}|{sat_det['UN_Com']}|{sat_det['Vlr_Item']}|{sat_det['CST_ICMS']}|{sat_det['CFOP']}|"
                            f.write(c810 + '\n')
                            clintot += 1

                    # C850: RESUMO DIÁRIO DE CF-E-SAT POR EQUIPAMENTO
                    rsSATResumo = self.db_conn.fetch_all(f"SELECT tbVendasDet.CST_ICMS, tbVendasDet.CFOP, tbVendasDet.Aliq_ICMS AS Aliq_ICMS, Sum(tbVendasDet.BaseCalculo) AS BaseCalculo,SUM(tbVendasDet.ValorTot) as ValorTot, Sum(tbVendasDet.Valor_ICMS) AS Valor_ICMS FROM (tbVendasDet INNER JOIN tbVendas ON tbVendasDet.IDVenda = tbVendas.ID) INNER JOIN tbVendasSAT ON tbVendas.ChaveNF = tbVendasSAT.ChaveCF WHERE tbVendas.ChaveNF = '{sat['ChaveCF']}' GROUP BY tbVendasDet.CST_ICMS, tbVendasDet.CFOP, tbVendasDet.Aliq_ICMS;")
                    if rsSATResumo:
                        for sat_resumo in rsSATResumo:
                            c850 = f"|C850|{sat_resumo['CST_ICMS']}|{sat_resumo['CFOP']}|{sat_resumo['Aliq_ICMS']}|{sat_resumo['ValorTot']}|{sat_resumo['BaseCalculo']}|{sat_resumo['Valor_ICMS']}||"
                            f.write(c850 + '\n')
                            clintot += 1

            # C990: ENCERRAMENTO DO BLOCO C
            cC990 = f"|C990|{clintot + 1}|"
            f.write(cC990 + '\n')
            clintot += 1

            # Bloco D - Documentos Fiscais II (Serviços ICMS)
            # ... (Lógica será adicionada em etapas)

            # Bloco E - Apuração do ICMS e IPI
            # ... (Lógica será adicionada em etapas)

            # 0990: ENCERRAMENTO DO BLOCO 0
            c0990 = f"|0990|{clintot + 1}|"
            f.write(c0990 + '\n')
            clintot += 1

            # B990: ENCERRAMENTO DO BLOCO B
            cB990 = f"|B990|{clintot + 1}|"
            f.write(cB990 + '\n')
            clintot += 1

            print(f"Arquivo EFD gerado em: {output_filename}")

        self.db_conn.disconnect_from_database()


# Exemplo de uso (para testes)
if __name__ == "__main__":
    generator = EFDGenerator()
    # Use datas de exemplo e um diretório temporário
    generator.generate_efd_icms_ipi("2023-01-01", "2023-01-31", "/tmp", "2023-01-01", "1")



