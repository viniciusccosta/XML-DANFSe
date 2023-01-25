# ======================================================================================================================
import xmltodict
import jinja2

from tkinter        import filedialog
from os             import path
from datetime       import datetime

# ======================================================================================================================
class NFSe:
    """
        6: <label id="serie_documento">
        7: <label id="numero_nf">

        10: <label id="prestador_razao_social">
        11: <label id="prestador_nome_fantasia">
        12: <label id="prestador_endereco">
        13: <label id="prestador_cep">
        14: <label id="prestador_telefone">
        15: <label id="prestador_email">
        16: <label id="prestador_inscricao_municipal">
        17: <label id="prestador_cnpj">

        18: <div id="data_geracao_texto">
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
        self.serie_documento                    = "Nota Fiscal de Serviço Eletrônica - NFS-e"                                   # TODO:
        self.numero_nf                          = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('Numero', )

        # Prestador:
        self.prestador_razao_social             = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('RazaoSocial', )
        self.prestador_nome_fantasia            = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('NomeFantasia', )
        self.prestador_endereco                 = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Endereco', {}).get('Endereco', {})       # TODO: Juntar as outras informações do endereço
        self.prestador_cep                      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Endereco', {}).get('Cep', )
        self.prestador_telefone                 = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Contato', {}).get('Telefone', )
        self.prestador_email                    = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('PrestadorServico', {}).get('Contato', {}).get('Email', )
        self.prestador_inscricao_municipal      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Prestador', {}).get('InscricaoMunicipal', )
        self.prestador_cnpj                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Prestador', {}).get('CpfCnpj', {}).get('Cnpj', )
        self.data_geracao_texto                 = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DataEmissao', )
        self.data_competencia                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Competencia', )
        self.cod_autencidade                    = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('CodigoVerificacao', )
        self.reponsavel_retencao                = None          # TODO:
        # self.qrcode                             =

        # Identificação:
        self.natureza_operacao                  = "Exigível"    # TODO: Seria "ExigibilidadeISS" ?
        self.numero_rps                         = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Rps', {}).get('IdentificacaoRps', {}).get('Numero', )
        self.serie_rps                          = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Rps', {}).get('IdentificacaoRps', {}).get('Serie', {})    # TODO: Converter número para "RPS - Recibo Provisórios de Serviços"
        self.data_emissao_rps                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Rps', {}).get('DataEmissao', )
        self.local_servicos                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoMunicipio', {})          # TODO: Tabela IBGE
        self.municipio_incidencia               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('MunicipioIncidencia', )

        # Tomador:
        self.tomador_cnpj                       = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('IdentificacaoTomador', {}).get('CpfCnpj', {}).get('Cnpj', )
        self.tomador_im                         = None          # TODO: .get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('', )
        self.tomador_razao_social               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('RazaoSocial', )
        self.tomador_endereco                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Endereco', )
        self.tomador_numero                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Numero', )
        self.tomador_complemento                = None          # TODO: xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}) ??
        self.tomador_bairro                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Bairro', )
        self.tomador_cep                        = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('Cep', )
        self.tomador_cidade                     = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Endereco', {}).get('CodigoMunicipio', {})   # TODO: Tabela IGBE
        self.tomador_telefone                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Contato', {}).get('Telefone')       # TODO: Telefone ?
        self.tomador_email                      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('TomadorServico', {}).get('Contato', {}).get('Email', )

        # Intermediário:
        self.intermediario_cnpj                 = None      # TODO: xml_dict.get('Nfse', {}).get('InfNfse', )
        self.intermediario_inscricao_municipal  = None      # TODO: xml_dict.get('Nfse', {}).get('InfNfse', )
        self.intermediario_razao_social         = None      # TODO: xml_dict.get('Nfse', {}).get('InfNfse', )

        # Descrição:
        self.descricao_servico                  = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Discriminacao', )

        # Tributos:
        self.atividade_municipio                = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DescricaoCodigoTributacaoMunicípio', )
        self.aliquota                           = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('ValoresNfse', {}).get('Aliquota', )
        self.item_lc116                         = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoTributacaoMunicipio', )
        self.cod_nbs                            = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoNbs', )
        self.cod_cnae                           = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('CodigoCnae', )
        self.valor_total                        = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorServicos', )
        self.desconto_incondicionado            = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('DescontoIncondicionado', )
        self.deducao_base_calculo               = None  # TODO: xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', )
        self.base_calculo                       = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('ValoresNfse', {}).get('BaseCalculo', )
        self.total_issqn                        = None  # TODO: xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', )
        self.issqn_retido                       = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('IssRetido', )
        self.desconto_condicionado              = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('DescontoCondicionado', )
        self.pis                                = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorPis', )
        self.cofins                             = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorCofins', )
        self.inss                               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorInss', )
        self.irrf                               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorIr', )
        self.csll                               = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('ValorCsll', )
        self.outras_retencoes                   = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', {}).get('OutrasRetencoes', )
        self.valor_issqn                        = None      # TODO: xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('Servico', {}).get('Valores', )
        self.valor_liquido                      = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('ValoresNfse', {}).get('ValorLiquidoNfse', )
        self.construcao_civil                   = None      # TODO: xml_dict.get('Nfse', {}).get('InfNfse', )
        self.cod_obra                           = None      # TODO: xml_dict.get('Nfse', {}).get('InfNfse', )
        self.art                                = None      # TODO: xml_dict.get('Nfse', {}).get('InfNfse', )

        # Informações:
        self.informacoes_adicionais             = xml_dict.get('Nfse', {}).get('InfNfse', {}).get('DeclaracaoPrestacaoServico', {}).get('InfDeclaracaoPrestacaoServico', {}).get('InformacoesComplementares', )

    def get_formated(self):
        data = {
            "serie_documento"                   : self.serie_documento,
            "numero_nf"                         : self.numero_nf,

            # Prestador:
            "prestador_razao_social"            : self.prestador_razao_social,
            "prestador_nome_fantasia"           : self.prestador_nome_fantasia,
            "prestador_endereco"                : self.prestador_endereco,
            "prestador_cep"                     : self.prestador_cep,
            "prestador_telefone"                : self.prestador_telefone,
            "prestador_email"                   : self.prestador_email,
            "prestador_inscricao_municipal"     : self.prestador_inscricao_municipal,
            "prestador_cnpj"                    : self.prestador_cnpj,
            "data_geracao_texto"                : self.data_geracao_texto,
            "data_competencia"                  : self.data_competencia,
            "cod_autencidade"                   : self.cod_autencidade,
            "reponsavel_retencao"               : self.reponsavel_retencao,
            # "qrcode"                            : self.qrcode,

            # Identificação:
            "natureza_operacao"                 : self.natureza_operacao,
            "numero_rps"                        : self.numero_rps,
            "serie_rps"                         : self.serie_rps,
            "data_emissao_rps"                  : self.data_emissao_rps,
            "local_servicos"                    : self.local_servicos,
            "municipio_incidencia"              : self.municipio_incidencia,

            # Tomador:
            "tomador_cnpj"                      : self.tomador_cnpj,
            "tomador_im"                        : self.tomador_im,
            "tomador_razao_social"              : self.tomador_razao_social,
            "tomador_endereco"                  : self.tomador_endereco,
            "tomador_numero"                    : self.tomador_numero,
            "tomador_complemento"               : self.tomador_complemento,
            "tomador_bairro"                    : self.tomador_bairro,
            "tomador_cep"                       : self.tomador_cep,
            "tomador_cidade"                    : self.tomador_cidade,
            "tomador_telefone"                  : self.tomador_telefone,
            "tomador_email"                     : self.tomador_email,

            # Intermediário:
            "intermediario_cnpj"                : self.intermediario_cnpj,
            "intermediario_inscricao_municipal" : self.intermediario_inscricao_municipal,
            "intermediario_razao_social"        : self.intermediario_razao_social,

            # Descrição:
            "descricao_servico"                 : self.descricao_servico,

            # Tributos:
            "atividade_municipio"               : self.atividade_municipio,
            "aliquota"                          : self.aliquota,
            "item_lc116"                        : self.item_lc116,
            "cod_nbs"                           : self.cod_nbs,
            "cod_cnae"                          : self.cod_cnae,
            "valor_total"                       : self.valor_total,
            "desconto_incondicionado"           : self.desconto_incondicionado,
            "deducao_base_calculo"              : self.deducao_base_calculo,
            "base_calculo"                      : self.base_calculo,
            "total_issqn"                       : self.total_issqn,
            "issqn_retido"                      : self.issqn_retido,
            "desconto_condicionado"             : self.desconto_condicionado,
            "pis"                               : self.pis,
            "cofins"                            : self.cofins,
            "inss"                              : self.inss,
            "irrf"                              : self.irrf,
            "csll"                              : self.csll,
            "outras_retencoes"                  : self.outras_retencoes,
            "valor_issqn"                       : self.valor_issqn,
            "valor_liquido"                     : self.valor_liquido,
            "construcao_civil"                  : self.construcao_civil,
            "cod_obra"                          : self.cod_obra,
            "art"                               : self.art,

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
    html = env.get_template('NFSe.html').render(nfse.get_formated())
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
        xml_dict = read_xml(file)
        nfse     = NFSe(xml_dict)
        html     = render_html(nfse)

        save_html(html, file)

# ======================================================================================================================
if __name__ == "__main__":
    main()

# ======================================================================================================================