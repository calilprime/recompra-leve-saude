# -*- coding: utf-8 -*-
"""
Caminhos, constantes e parâmetros fixos da automação de recompra.
=================================================================

Regra da casa: nada de juros, multa ou limite de linha escrito no código.
O que está aqui é estrutura (nome de coluna, nome de aba, endereço de célula),
não parâmetro financeiro — esses vêm sempre da aba INFORMAÇÕES.
"""

from decimal import Decimal
from pathlib import Path

# ---------------------------------------------------------------------------
#  Pastas do projeto
# ---------------------------------------------------------------------------
RAIZ = Path(__file__).resolve().parent.parent

PASTA_TEMPLATES = RAIZ / "templates"
PASTA_SNAPSHOTS = RAIZ / "snapshots"
PASTA_SAIDAS = RAIZ / "saidas"
PASTA_LOGS = RAIZ / "logs"
PASTA_AMOSTRAS = RAIZ / "testes" / "amostras"

ARQUIVO_HISTORICO = RAIZ / "historico_recompras.json"
TEMPLATE_PADRAO = PASTA_TEMPLATES / "Leve_Saude_Recompra_TEMPLATE.xlsx"

# Pasta oficial das rodadas anteriores (OneDrive) — SOMENTE LEITURA.
PASTA_ONEDRIVE = Path(
    r"C:\Users\Carlos Fontes\NETZ ASSET\CRÉDITO - General"
    r"\04.Cobranca\10.FIDC_Leve_Saude\03.Recompra"
)

for _p in (PASTA_TEMPLATES, PASTA_SNAPSHOTS, PASTA_SAIDAS, PASTA_LOGS):
    _p.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
#  Status da Grafeno e prioridade de desempate (decisão 1.2)
# ---------------------------------------------------------------------------
STATUS_RECOMPRAVEL = "Aberta (Vencida)"

#  Quando a mesma chave tem mais de um boleto, vence o primeiro desta lista.
#  Nunca "o primeiro da planilha" — foi o defeito 2.2.
PRIORIDADE_STATUS = ["Aberta (Vencida)", "Aberta", "Baixada", "Paga"]

#  Rótulos aceitos vindos da API (preenchido na Fase 5, se a API devolver códigos).
MAPA_STATUS_API = {}


# ---------------------------------------------------------------------------
#  Casamento (decisão 1.3 + seção 3.1)
# ---------------------------------------------------------------------------
#  Tolerância de valor: só vale quando o par (documento, vencimento) tem
#  exatamente uma linha não casada de cada lado. Em 192 de 227 casos o sacado
#  tem mais de um título no mesmo vencimento — relaxar isso pareia o errado.
TOLERANCIA_VALOR = Decimal("0.05")


# ---------------------------------------------------------------------------
#  Ligação direta título -> boleto (verificação de identificadores, Fase 3)
# ---------------------------------------------------------------------------
#  O boleto que a Grafeno emite para um título carrega o NumeroTitulo da Vórtx
#  dentro do Nosso_Número, com um prefixo de convênio:
#
#      NumeroTitulo 382477119  ->  Nosso_Número 403382477119
#
#  Medido no arquivo da 12ª: 62.698 boletos seguem esse padrão.
#
#  **Isto não substitui a chave composta, e a distinção é cara.** Usada
#  sozinha, a ligação encontra 2.493 títulos a mais — e em 79% deles o boleto é
#  de OUTRA pessoa. A numeração é sequencial e densa dos dois lados (distância
#  típica de 8 entre vizinhos), então, quando o título não tem boleto,
#  "403"+NumeroTitulo cai no boleto do vizinho.
#
#  O uso legítimo é estreito: **desempatar entre os boletos que a chave
#  composta já validou** por documento, vencimento e valor. Aí ela diz qual dos
#  candidatos é de fato daquele título, em vez de escolher pela prioridade de
#  status. Nunca cria par novo.
PREFIXO_NOSSO_NUMERO = "403"


# ---------------------------------------------------------------------------
#  Custo da defasagem entre extração e recompra (seção 2.4)
# ---------------------------------------------------------------------------
#  Medido entre a v1.1 e a v1.5 da 12ª: 74 títulos e R$ 89.640,46 em 4 dias úteis.
TITULOS_POR_DIA_UTIL = 19          # faixa observada: 18 a 20
VALOR_POR_DIA_UTIL = Decimal("22000")


# ---------------------------------------------------------------------------
#  Fundo
# ---------------------------------------------------------------------------
NOME_FUNDO = "FIDC LEVE SAÚDE"

#  Descoberto na Fase 3, com SELECT DISTINCT "nomeFundo","cnpjFundo".
#  Nome completo no banco:
#    NTZ LEVE SAUDE FUNDO DE INVESTIMENTO EM DIREITOS CREDITÓRIOS COMERCIAIS
#    _ RESPONSABILIDADE LIMITADA
#
#  O filtro não é opcional: a tabela guarda mais de um fundo. Na posição de
#  12/08/2026 são 57.750 linhas do Leve Saúde e 156 do NTZ IMPULSE
#  (62.122.004/0001-08) — sacados pessoa jurídica, valores na casa dos R$ 200
#  mil, que não têm nada a ver com esta recompra.
CNPJ_FUNDO = "61.302.922/0001-48"


# ---------------------------------------------------------------------------
#  Abas do arquivo de recompra
# ---------------------------------------------------------------------------
ABA_INFORMACOES = "INFORMAÇÕES"
ABA_TERMO = "TERMO DE RECOMPRA"
ABA_GRAFENO = "GRAFENO"
ABA_VORTX = "VORTX"


# ---------------------------------------------------------------------------
#  Aba INFORMAÇÕES — endereços das células de parâmetro
# ---------------------------------------------------------------------------
CEL_DATA_RECOMPRA = "D8"
CEL_NUMERO_RODADA = "D9"
CEL_JANELA_INICIO = "D10"
CEL_JANELA_FIM = "D11"
CEL_QTD_RECOMPRA = "D12"
CEL_JUROS_MORA = "D13"
CEL_MULTA = "D14"
CEL_VALOR_TOTAL = "C19"


# ---------------------------------------------------------------------------
#  Aba VORTX — colunas de dado (A..AM) e colunas auxiliares (AP..AZ)
# ---------------------------------------------------------------------------
#  Ordem exata do export da Vórtx. É por esta lista que a aba é reescrita:
#  posição por posição, sem criar, mover ou renomear coluna nenhuma.
COLUNAS_VORTX = [
    "Situacao", "PES_TIPO_PESSOA", "CedenteCnpjCpf", "TIT_CEDENTE_ENT_CODIGO",
    "CedenteNome", "Cnae", "SecaoCNAEDescricao", "NotaPdd", "SAC_TIPO_PESSOA",
    "SacadoCnpjCpf", "SacadoNome", "IdTituloVortx", "TipoAtivo", "DataEmissao",
    "DataAquisicao", "DataVencimento", "NumeroBoleto", "NumeroTitulo",
    "CampoChave", "ValorAquisicao", "ValorNominal", "ValorPresente", "PDDNota",
    "PDDVencido", "PagamentoParcial", "Coobricacao", "DataGeracao", "PDDTotal",
    "CampoAdicional1", "CampoAdicional2", "CampoAdicional3", "CampoAdicional4",
    "CampoAdicional5", "PDDEfeitoVagao", "PercentagemEfeitoVagao",
    "IdTituloVortxOriginador", "Registradora", "IdContratoRegistradora",
    "IdTituloRegistradora",
]

#  Colunas que o motor lê de fato.
VORTX_DOCUMENTO = "SacadoCnpjCpf"
VORTX_NOME = "SacadoNome"
VORTX_ID = "IdTituloVortx"
VORTX_VENCIMENTO = "DataVencimento"
VORTX_VALOR = "ValorNominal"
VORTX_GERACAO = "DataGeracao"
VORTX_EMISSAO = "DataEmissao"

#  Colunas auxiliares da planilha (a automação as reescreve como fórmula).
#  Letras das colunas de dado que as fórmulas citam. Saem da ordem de
#  COLUNAS_VORTX: NumeroTitulo é a 18ª, logo R.
VORTX_COL_NUMERO_TITULO = "R"

VORTX_COL_AUX = "AP"          # chave concatenada
VORTX_COL_CHECK = "AQ"        # quantos boletos a chave tem na Grafeno
VORTX_COL_STATUS = "AR"       # status resolvido
VORTX_COL_ELEGIVEL = "AS"     # 1 se entra no Termo
VORTX_COL_VENCIMENTO = "AT"
VORTX_COL_RECOMPRA = "AU"
VORTX_COL_VALOR = "AV"
VORTX_COL_PRAZO = "AW"
VORTX_COL_MULTA = "AX"
VORTX_COL_JUROS = "AY"
VORTX_COL_TOTAL = "AZ"


# ---------------------------------------------------------------------------
#  Aba GRAFENO — colunas do export de Cobranças
# ---------------------------------------------------------------------------
COLUNAS_GRAFENO = [
    "Usuário_Criador_da_Cobrança", "Conta_Origem", "Pagador", "Número_Documento",
    "Seu_Número", "Régua_Cobrança", "Data_Criação", "Data_Vencimento",
    "Valor_Cobrança", "Grupo_Cobrança", "Multa_Atraso", "Modelos_Juros",
    "Juros_Mês", "Desconto", "Desconto_Permitido_Até", "Abatimento",
    "Nosso_Número", "Data_Pagamento", "Valor_Pago", "Status", "Status_Protesto",
    "Registro_Boleto", "Banco_de_Origem", "Agência_de_Origem", "Conta_de_Origem",
    "Nome_do_Pagador", "Documento", "Forma_de_Pagamento",
]

GRAFENO_DOCUMENTO = "Número_Documento"
GRAFENO_NOME = "Pagador"
GRAFENO_VENCIMENTO = "Data_Vencimento"
GRAFENO_VALOR = "Valor_Cobrança"
GRAFENO_STATUS = "Status"
GRAFENO_NOSSO_NUMERO = "Nosso_Número"
GRAFENO_SEU_NUMERO = "Seu_Número"
GRAFENO_DATA_PAGAMENTO = "Data_Pagamento"
GRAFENO_VALOR_PAGO = "Valor_Pago"

#  Idem do lado da Grafeno: Nosso_Número é a 17ª coluna (Q) e Status a 20ª (T).
GRAFENO_COL_NOSSO = "Q"
GRAFENO_COL_STATUS = "T"

GRAFENO_COL_AUX = "AE"        # chave concatenada, coluna auxiliar da planilha


# ---------------------------------------------------------------------------
#  Termo de recompra — layout da aba
# ---------------------------------------------------------------------------
TERMO_PRIMEIRA_LINHA = 5
TERMO_COLUNAS = {          # coluna da planilha -> campo do motor
    "B": "nome",
    "C": "documento_formatado",
    "D": "id_titulo",
    "E": "vencimento",
    "F": "data_recompra",
    "G": "valor_nominal",
    "H": "valor_recompra",
}

#  Regra de nome do arquivo, conforme INFORMAÇÕES!C6, com a vírgula decimal que
#  as rodadas anteriores usam ("Leve_Saude_12_Recompra_0,59MM_v1.5_08-2026").
PADRAO_NOME_SAIDA = "Leve_Saude_{n}_Recompra_{valor}MM_v1.0_{mes_ano}.xlsx"


# ---------------------------------------------------------------------------
#  Causas de exceção (passada 3 da conciliação)
# ---------------------------------------------------------------------------
CAUSA_SACADO_INEXISTENTE = "SACADO_INEXISTENTE"
CAUSA_DIVERGENCIA_VALOR = "DIVERGENCIA_VALOR"
CAUSA_TITULO_IRMAO = "TITULO_IRMAO"
CAUSA_DIVERGENCIA_VENCIMENTO = "DIVERGENCIA_VENCIMENTO"
CAUSA_SEM_CORRESPONDENCIA = "SEM_CORRESPONDENCIA"
