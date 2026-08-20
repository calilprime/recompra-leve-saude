# -*- coding: utf-8 -*-
"""
Fonte Grafeno — estoque de boletos.
===================================

Primária (Fase 5): API de Cobranças, ``https://docs.grafeno.digital/v1.0``,
**somente leitura**. Enquanto a conta não sai, o caminho é o arquivo exportado
do portal, que precisa funcionar por completo.

Escopo travado: nada que emita, cancele, baixe ou liquide boleto pode existir
neste módulo. Só listagem.
"""

import os
from pathlib import Path

from core import config
from core import leitura
from core import normalizacao as nz
from core.fonte_vortx import ResultadoFonte


class Boleto:
    """Uma linha do export de cobranças da Grafeno, já normalizada."""

    __slots__ = (
        "linha", "documento", "nome", "vencimento", "valor", "status",
        "nosso_numero", "seu_numero", "data_pagamento", "valor_pago",
        "data_criacao", "chave", "chave_parcial", "chave_boleto", "bruto",
    )

    def __init__(self, **campos):
        for nome in self.__slots__:
            setattr(self, nome, campos.get(nome))

    @property
    def prioridade(self):
        """Posição do status na ordem de desempate da decisão 1.2."""
        try:
            return config.PRIORIDADE_STATUS.index(self.status)
        except ValueError:
            return len(config.PRIORIDADE_STATUS)

    def __repr__(self):
        return (f"<Boleto {self.documento} venc={self.vencimento} "
                f"valor={self.valor} status={self.status!r}>")


# ---------------------------------------------------------------------------
#  Arquivo
# ---------------------------------------------------------------------------
def ler_arquivo(caminho, aba=config.ABA_GRAFENO, log=None):
    """Lê o estoque de boletos de um arquivo (.xlsx com aba, ou .csv)."""
    log = log or (lambda *a, **k: None)
    caminho = Path(caminho)
    log(f"  Lendo estoque Grafeno de {caminho.name}…")

    tabela = leitura.ler(
        caminho, aba=aba,
        colunas_sentinela=(config.GRAFENO_DOCUMENTO,),
        log=log,
    )
    return _montar(tabela, origem="arquivo", caminho=caminho, log=log)


def _montar(tabela, origem, caminho, log=None):
    log = log or (lambda *a, **k: None)

    p_doc = tabela.posicao(config.GRAFENO_DOCUMENTO)
    p_nome = tabela.posicao(config.GRAFENO_NOME, obrigatoria=False)
    p_venc = tabela.posicao(config.GRAFENO_VENCIMENTO)
    p_valor = tabela.posicao(config.GRAFENO_VALOR)
    p_status = tabela.posicao(config.GRAFENO_STATUS)
    p_nosso = tabela.posicao(config.GRAFENO_NOSSO_NUMERO, obrigatoria=False)
    p_seu = tabela.posicao(config.GRAFENO_SEU_NUMERO, obrigatoria=False)
    p_pgto = tabela.posicao(config.GRAFENO_DATA_PAGAMENTO, obrigatoria=False)
    p_valor_pago = tabela.posicao(config.GRAFENO_VALOR_PAGO, obrigatoria=False)
    p_criacao = tabela.posicao("Data_Criação", obrigatoria=False)

    def pega(linha, posicao):
        if posicao is None or posicao >= len(linha):
            return None
        return linha[posicao]

    #  Mesma ideia do lado da Vórtx: a linha que volta para a aba GRAFENO é
    #  montada por nome, na ordem do export, sem a coluna auxiliar AE.
    posicoes_saida = [tabela.posicao(nome, obrigatoria=False)
                      for nome in config.COLUNAS_GRAFENO]

    registros = []
    status_desconhecidos = {}

    for numero, linha in tabela.linhas:
        documento = nz.documento(pega(linha, p_doc))
        vencimento = nz.data(pega(linha, p_venc))
        valor = nz.valor(pega(linha, p_valor))
        status = nz.texto(pega(linha, p_status))
        status = config.MAPA_STATUS_API.get(status, status)
        if status and status not in config.PRIORIDADE_STATUS:
            status_desconhecidos[status] = status_desconhecidos.get(status, 0) + 1

        nosso_numero = nz.texto(pega(linha, p_nosso))
        registros.append(Boleto(
            linha=numero,
            documento=documento,
            nome=nz.texto(pega(linha, p_nome)),
            vencimento=vencimento,
            valor=valor,
            status=status,
            nosso_numero=nosso_numero,
            seu_numero=nz.texto(pega(linha, p_seu)),
            data_pagamento=nz.data(pega(linha, p_pgto)),
            valor_pago=nz.valor(pega(linha, p_valor_pago)),
            data_criacao=nz.data(pega(linha, p_criacao)),
            chave=nz.chave(documento, vencimento, valor),
            chave_parcial=nz.chave_parcial(documento, vencimento),
            #  Só os dígitos do Nosso_Número: é contra ele que o título compara
            #  o seu "403"+NumeroTitulo na hora de desempatar.
            chave_boleto=nz.so_digitos(nosso_numero),
            bruto=tuple(pega(linha, p) for p in posicoes_saida),
        ))

    avisos = []
    if status_desconhecidos:
        avisos.append(
            "Status fora da lista de prioridade da decisão 1.2: "
            + ", ".join(f"{s!r} ({n})" for s, n in sorted(status_desconhecidos.items()))
        )

    #  ------------------------------------------------------------------
    #  Até quando a base vai, deduzido do conteúdo.
    #
    #  O export da Grafeno não traz a data da extração em coluna nenhuma, e a
    #  data de modificação do arquivo mente: o arquivo lido em 14/08 estava
    #  gravado naquele dia e o conteúdo ia só até 08/08. A defasagem saiu como
    #  "1 dia útil" quando era de seis, e a rodada passou.
    #
    #  ``Data_Criação`` é quando o boleto foi emitido e ``Data_Pagamento``
    #  quando foi pago; nenhuma das duas pode ser posterior à extração, então a
    #  maior das duas é o piso da data de extração. É estimativa por baixo, e é
    #  rotulada assim em toda parte — mas é uma estimativa que não mente na
    #  direção perigosa: nunca faz a base parecer mais nova do que é.
    marcos = {}
    maior_criacao = max((b.data_criacao for b in registros if b.data_criacao),
                        default=None)
    maior_pagamento = max((b.data_pagamento for b in registros if b.data_pagamento),
                          default=None)
    if maior_criacao:
        marcos["Data_Criação"] = maior_criacao
    if maior_pagamento:
        marcos["Data_Pagamento"] = maior_pagamento
    data_conteudo = max(marcos.values()) if marcos else None

    log(f"    {len(registros)} boletos normalizados.")
    if data_conteudo:
        log(f"    Conteúdo da base vai até {nz.br(data_conteudo)} "
            + " · ".join(f"MAX({c})={nz.br(d)}" for c, d in sorted(marcos.items()))
            + ".")
    else:
        avisos.append(
            "Não há Data_Criação nem Data_Pagamento preenchidas na base da "
            "Grafeno — sem elas não dá para saber de que dia é a extração."
        )
    return ResultadoFonte(registros, tabela, origem, caminho,
                          datas_geracao=None, avisos=avisos,
                          data_conteudo=data_conteudo, marcos_conteudo=marcos)


# ---------------------------------------------------------------------------
#  API de Cobranças (Fase 5 — depende da conta e do token)
# ---------------------------------------------------------------------------
BASE_API = "https://api.grafeno.digital"
DOC_API = "https://docs.grafeno.digital/v1.0"


def credenciais_api(env=None):
    """Token e número da conta, lidos do .env."""
    env = env or os.environ
    from core.fonte_vortx import _ler_env_arquivo
    valores = dict(_ler_env_arquivo())
    valores.update(env)
    token = valores.get("GRAFENO_TOKEN") or ""
    conta = valores.get("GRAFENO_ACCOUNT_NUMBER") or ""
    return token, conta


def api_disponivel():
    """True quando há token e conta configurados."""
    token, conta = credenciais_api()
    return bool(token and conta)


def ler_api(*args, **kwargs):
    """
    Extração via API de Cobranças.

    Ainda não implementada: a conta Grafeno foi solicitada em 11/08 e o token
    não existe (pendência 14.2-3). O endpoint, o parâmetro de filtro por período
    e o formato de paginação precisam ser lidos da documentação antes — a regra
    15.10 proíbe inventar nome de rota ou de campo.

    Quando for implementada, duas travas são obrigatórias:
      1. comparar o total informado pela API com o total de registros recebidos
         e abortar em caso de divergência (truncamento silencioso foi o defeito 2.1);
      2. nenhuma chamada de escrita — emitir, cancelar, baixar ou liquidar.
    """
    raise NotImplementedError(
        "A API da Grafeno ainda não foi liberada (pendência 14.2-3). "
        "Use a opção 'Arquivo' na interface."
    )
