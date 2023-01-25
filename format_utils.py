from re             import findall
from babel.numbers  import format_currency

# ======================================================================================================================
def formatar_cep(cep: str):
    digitos = ''.join(findall(r"\d", cep)).zfill(8)
    return f'{digitos[0:0+2]}.{digitos[2:2+3]}-{digitos[5:5+3]}'

def formatar_cpf(cpf: str):
    digitos = ''.join(findall(r"\d", cpf)).zfill(11)
    return f'{digitos[0:0+3]}.{digitos[3:3+3]}.{digitos[6:6+3]}-{digitos[9:9+2]}'

def formatar_cnpj(cnpj: str):
    digitos = ''.join(findall(r"\d", cnpj)).zfill(14)
    return f'{digitos[0:0+2]}.{digitos[2:2+3]}.{digitos[5:5+3]}/{digitos[8:8+4]}-{digitos[12:12+2]}'

def formatar_moeda(valor: str):
    valor = valor.zfill(0)
    return format_currency(valor, currency="BRL", locale="pt_BR")

# ======================================================================================================================
