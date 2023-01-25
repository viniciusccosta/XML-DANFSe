# ======================================================================================================================
import xmltodict
import jinja2

from tkinter        import filedialog
from os             import path
from datetime       import datetime
from dateutil       import parser
from babel.numbers  import format_decimal

from format_utils   import *

# ======================================================================================================================
class NFSe:
    """
        6: <label id="tipo_rps">
        7: <label id="numero_nf">

        10: <label id="prestador_razao_social">
        11: <label id="prestador_nome_fantasia">
        12: <label id="prestador_endereco">
        13: <label id="prestador_cep">
        14: <label id="prestador_telefone">
        15: <label id="prestador_email">
        16: <label id="prestador_inscricao_municipal">
        17: <label id="prestador_cnpj">

        18: <div id="data_geracao">
        19: <div id="data_competencia">
        20: <div id="cod_autencidade">
        21: <div id="reponsavel_retencao">
        22: <div id="qrcode_container" class="stack">

        23: <img id="qrcode" src="" alt="QR CODE">

        25: <label id="natureza_operacao">
        26: <label id="numero_rps">
        27: <label id="serie_rps">
        28: <label id="data_emissao_rps">
        29: <label id="local_servicos">
        30: <label id="municipio_incidencia">
        31: <div id="tomador" class="container stack">
        32: <label id="tomador_cnpj">
        33: <label id="tomador_im">
        34: <label id="tomador_razao_social">
        35: <label id="tomador_endereco">
        36: <label id="tomador_numero">
        37: <label id="tomador_complemento">
        38: <label id="tomador_bairro">
        39: <label id="tomador_cep">
        40: <label id="tomador_cidade">
        41: <label id="tomador_telefone">
        42: <label id="tomador_email">
        43: <div id="intermediario" class="container stack">
        44: <label id="intermediario_cnpj">
        45: <label id="intermediario_inscricao_municipal">
        46: <label id="intermediario_razao_social">
        47: <div id="servicos" class="container stack">
        48: <div id="descricao_servico">
        49: <div id="tributos" class="container stack">
        50: <div id="atividade_municipio">
        51: <div id="aliquota">
        52: <div id="item_lc116">
        53: <div id="cod_nbs">
        54: <div id="cod_cnae">
        55: <div id="valor_total">
        56: <div id="desconto_incondicionado">
        57: <div id="deducao_base_calculo">
        58: <div id="base_calculo">
        59: <div id="total_issqn">
        60: <div id="issqn_retido">
        61: <div id="desconto_condicionado">
        62: <div id="pis">
        63: <div id="cofins">
        64: <div id="inss">
        65: <div id="irrf">
        66: <div id="csll">
        67: <div id="outras_retencoes">
        68: <div id="valor_issqn">
        69: <div id="valor_liquido">
        70: <div id="construcao_civil">
        71: <div id="cod_obra">
        72: <div id="art">
        74: <div id="informacoes_adicionais">
    """
    def __init__(self, xml_dict: dict):
        # Cabeçalho
        self.tipo_rps                           = "Nota Fiscal de Serviço Eletrônica - NFS-e"                           # TODO:
        self.numero_nf                          = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('Numero', None)

        # Prestador:
        self.prestador_razao_social             = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('RazaoSocial', None)
        self.prestador_nome_fantasia            = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('NomeFantasia', None)
        self.prestador_endereco                 = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Endereco', {}).get('Endereco', {})       # TODO: Juntar as outras informações do endereço
        self.prestador_cep                      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Endereco', {}).get('Cep', None)
        self.prestador_telefone                 = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Contato', {}).get('Telefone', None)
        self.prestador_email                    = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Contato', {}).get('Email', None)
        self.prestador_inscricao_municipal      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Prestador', {}).get('InscricaoMunicipal', None)
        self.prestador_cnpj                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Prestador', {}).get('CpfCnpj', {}).get('Cnpj', None)
        self.prestador_cpf                      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Prestador', {}).get('CpfCnpj', {}).get('Cpf', None)
        self.data_geracao                       = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DataEmissao', None)
        self.data_competencia                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Competencia', None)
        self.cod_autencidade                    = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('CodigoVerificacao', None)
        self.reponsavel_retencao                = None          # TODO:
        self.qrcode                             = None          # TODO:

        # Identificação:
        self.natureza_operacao                  = None          # TODO: DeclaracaoPrestacaoServico > InfDeclaracaoPrestacaoServico > Servico > ExigibilidadeISS ?
        self.numero_rps                         = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Rps', {}).get('IdentificacaoRps', {}).get('Numero', None)
        self.serie_rps                          = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Rps', {}).get('IdentificacaoRps', {}).get('Serie', None)      # TODO: Converter número para "RPS - Recibo Provisórios de Serviços"
        self.data_emissao_rps                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Rps', {}).get('DataEmissao', None)
        self.local_servicos                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoMunicipio', None)                                        # TODO: Tabela IBGE
        self.municipio_incidencia               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('MunicipioIncidencia', None)                                    # TODO: Tabela IBGE

        # Tomador:
        self.tomador_cnpj                       = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('IdentificacaoTomador', {}).get('CpfCnpj', {}).get('Cnpj', None)
        self.tomador_cpf                        = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('IdentificacaoTomador', {}).get('CpfCnpj', {}).get('Cpf', None)
        self.tomador_im                         = None          # TODO:
        self.tomador_razao_social               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('RazaoSocial', None)
        self.tomador_endereco                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Endereco', None)
        self.tomador_numero                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Numero', None)
        self.tomador_complemento                = None          # TODO:
        self.tomador_bairro                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Bairro', None)
        self.tomador_cep                        = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Cep', None)
        self.tomador_cidade                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('CodigoMunicipio', None)             # TODO: Tabela IGBE
        self.tomador_telefone                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Contato', {}).get('Telefone', None)
        self.tomador_email                      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Contato', {}).get('Email', None)

        # Intermediário:
        self.intermediario_cnpj                 = None          # TODO:
        self.intermediario_inscricao_municipal  = None          # TODO:
        self.intermediario_razao_social         = None          # TODO:

        # Descrição:
        self.descricao_servico                  = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Discriminacao', None)

        # Tributos:
        self.cod_tributacao                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoTributacaoMunicipio', None)
        self.atividade_municipio                = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DescricaoCodigoTributacaoMunicípio', None)
        self.aliquota                           = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('ValoresNfse', {}).get('Aliquota', None)
        self.item_lc116                         = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoTributacaoMunicipio', None)
        self.cod_nbs                            = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoNbs', None)
        self.cod_cnae                           = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoCnae', None)
        self.valor_total                        = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorServicos', None)
        self.desconto_incondicionado            = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('DescontoIncondicionado', None)
        self.deducao_base_calculo               = None          # TODO:
        self.base_calculo                       = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('ValoresNfse', {}).get('BaseCalculo', None)
        self.total_issqn                        = None          # TODO: ValoresNfse > ValorIss ?
        self.issqn_retido                       = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('IssRetido', None)
        self.desconto_condicionado              = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('DescontoCondicionado', None)
        self.pis                                = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorPis', None)
        self.cofins                             = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorCofins', None)
        self.inss                               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorInss', None)
        self.irrf                               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorIr', None)
        self.csll                               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorCsll', None)
        self.outras_retencoes                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('OutrasRetencoes', None)
        self.valor_issqn                        = None          # TODO: ValoresNfse > ValorIss ?
        self.valor_liquido                      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('ValoresNfse', {}).get('ValorLiquidoNfse', None)
        self.construcao_civil                   = None          # TODO:
        self.cod_obra                           = None          # TODO:
        self.art                                = None          # TODO:

        # Informações:
        self.informacoes_adicionais             = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('InformacoesComplementares', None)

    def get_formated(self):
        data = {
            "tipo_rps"                          : self.tipo_rps if self.tipo_rps else "",
            "numero_nf"                         : self.numero_nf if self.numero_nf else "",

            # Prestador:
            "prestador_razao_social"            : self.prestador_razao_social if self.prestador_razao_social else "",
            "prestador_nome_fantasia"           : self.prestador_nome_fantasia if self.prestador_nome_fantasia else "",
            "prestador_endereco"                : self.prestador_endereco if self.prestador_endereco else "",
            "prestador_cep"                     : formatar_cep(self.prestador_cep) if self.prestador_cep else "",
            "prestador_telefone"                : self.prestador_telefone if self.prestador_telefone else "",
            "prestador_email"                   : self.prestador_email if self.prestador_email else "",
            "prestador_inscricao_municipal"     : self.prestador_inscricao_municipal if self.prestador_inscricao_municipal else "",
            "prestador_cnpj"                    : formatar_cnpj(self.prestador_cnpj) if self.prestador_cnpj else "",
            "prestador_cpf"                     : formatar_cpf(self.prestador_cpf) if self.prestador_cpf else "",
            "data_geracao"                      : parser.parse(self.data_geracao).strftime("%d/%m/%Y %H:%M:%S") if self.data_geracao else "",
            "data_competencia"                  : parser.parse(self.data_competencia).strftime("%d/%m/%Y") if self.data_competencia else "",
            "cod_autencidade"                   : self.cod_autencidade if self.cod_autencidade else "",
            "reponsavel_retencao"               : self.reponsavel_retencao if self.reponsavel_retencao else "",
            "qrcode"                            : self.qrcode if self.qrcode else "",

            # Identificação:
            "natureza_operacao"                 : self.natureza_operacao if self.natureza_operacao else "",
            "numero_rps"                        : self.numero_rps if self.numero_rps else "",
            "serie_rps"                         : self.serie_rps if self.serie_rps else "",
            "data_emissao_rps"                  : parser.parse(self.data_emissao_rps).strftime("%d/%m/%Y") if self.data_emissao_rps else "",
            "local_servicos"                    : self.local_servicos if self.local_servicos else "",
            "municipio_incidencia"              : self.municipio_incidencia if self.municipio_incidencia else "",

            # Tomador:
            "tomador_cnpj"                      : formatar_cnpj(self.tomador_cnpj) if self.tomador_cnpj else "",
            "tomador_cpf"                       : formatar_cpf(self.tomador_cpf) if self.tomador_cpf else "",
            "tomador_im"                        : self.tomador_im if self.tomador_im else "",
            "tomador_razao_social"              : self.tomador_razao_social if self.tomador_razao_social else "",
            "tomador_endereco"                  : self.tomador_endereco if self.tomador_endereco else "",
            "tomador_numero"                    : self.tomador_numero if self.tomador_numero else "",
            "tomador_complemento"               : self.tomador_complemento if self.tomador_complemento else "",
            "tomador_bairro"                    : self.tomador_bairro if self.tomador_bairro else "",
            "tomador_cep"                       : formatar_cep(self.tomador_cep) if self.tomador_cep else "",
            "tomador_cidade"                    : self.tomador_cidade if self.tomador_cidade else "",
            "tomador_telefone"                  : self.tomador_telefone if self.tomador_telefone else "",
            "tomador_email"                     : self.tomador_email if self.tomador_email else "",

            # Intermediário:
            "intermediario_cnpj"                : formatar_cnpj(self.intermediario_cnpj) if self.intermediario_cnpj else "",
            "intermediario_inscricao_municipal" : self.intermediario_inscricao_municipal if self.intermediario_inscricao_municipal else "",
            "intermediario_razao_social"        : self.intermediario_razao_social if self.intermediario_razao_social else "",

            # Descrição:
            "descricao_servico"                 : self.descricao_servico if self.descricao_servico else "",

            # Tributos:
            "cod_tributacao"                    : self.cod_tributacao if self.cod_tributacao else "",
            "atividade_municipio"               : self.atividade_municipio if self.atividade_municipio else "",
            "aliquota"                          : format_decimal(self.aliquota, format="#0.00", locale="pt_BR")if self.aliquota else "",
            "item_lc116"                        : self.item_lc116 if self.item_lc116 else "",
            "cod_nbs"                           : self.cod_nbs if self.cod_nbs else "",
            "cod_cnae"                          : self.cod_cnae if self.cod_cnae else "",

            "valor_total"                       : formatar_moeda(self.valor_total) if self.valor_total else "",
            "desconto_incondicionado"           : formatar_moeda(self.desconto_incondicionado) if self.desconto_incondicionado else "",
            "deducao_base_calculo"              : formatar_moeda(self.deducao_base_calculo) if self.deducao_base_calculo else "",
            "base_calculo"                      : formatar_moeda(self.base_calculo) if self.base_calculo else "",
            "total_issqn"                       : formatar_moeda(self.total_issqn) if self.total_issqn else "",
            "issqn_retido"                      : self.issqn_retido if self.issqn_retido else "",
            "desconto_condicionado"             : formatar_moeda(self.desconto_condicionado) if self.desconto_condicionado else "",

            "pis"                               : formatar_moeda(self.pis) if self.pis else "",
            "cofins"                            : formatar_moeda(self.cofins) if self.cofins else "",
            "inss"                              : formatar_moeda(self.inss) if self.inss else "",
            "irrf"                              : formatar_moeda(self.irrf) if self.irrf else "",
            "csll"                              : formatar_moeda(self.csll) if self.csll else "",
            "outras_retencoes"                  : formatar_moeda(self.outras_retencoes) if self.outras_retencoes else "",
            "valor_issqn"                       : formatar_moeda(self.valor_issqn) if self.valor_issqn else "",
            "valor_liquido"                     : formatar_moeda(self.valor_liquido) if self.valor_liquido else "",

            "construcao_civil"                  : self.construcao_civil if self.construcao_civil else "",
            "cod_obra"                          : self.cod_obra if self.cod_obra else "",
            "art"                               : self.art if self.art else "",

            # Informações Complementares
            "informacoes_adicionais"            : self.informacoes_adicionais if self.informacoes_adicionais else "",
        }

        return data

    def __str__(self):
        return str(self.__dict__)

# ======================================================================================================================
def read_xml(xmlfile: str):
    with open(xmlfile, 'r', encoding='utf8', ) as f:
        result = xmltodict.parse(f.read())
        return result

def render_html(nfse: NFSe,):
    env  = jinja2.Environment(loader = jinja2.FileSystemLoader('./'))
    html = env.get_template('templates/nfse.html').render(nfse.get_formated())
    return html

def save_html(html, xml_filename):
    xml_filename  = path.splitext(xml_filename)[0]
    agora         = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    html_filename = f'{xml_filename}-{agora}.html'

    with open(html_filename, 'w', encoding='utf8') as hf:
        hf.write(html)

def main():
    files = filedialog.askopenfilenames(filetypes=[('XML', '*.xml')])

    for file in files:
        try:
            xml_dict = read_xml(file)
            nfse     = NFSe(xml_dict)
            html     = render_html(nfse)

            save_html(html, file)
        except Exception as e:
            print(f"Erro ao gerar HTML do arquivo {file}: ", repr(e), )

# ======================================================================================================================
if __name__ == "__main__":
    main()

# ======================================================================================================================
