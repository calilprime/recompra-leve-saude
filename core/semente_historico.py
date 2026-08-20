# -*- coding: utf-8 -*-
"""
Semente do histórico a partir das rodadas já emitidas (item 5).
==============================================================

O ``historico_recompras.json`` estava vazio, e três coisas dependem dele:

* a **numeração** da próxima rodada — foi por isso que a simulação de 14/08 saiu
  como ``Leve_Saude_1_Recompra_...`` em vez de 14;
* a **trava contra recompra repetida**, que existe por causa do título
  165955407 (Cíntia Farias Cordeiro), cobrado três vezes entre duas rodadas;
* a **validação de sobreposição de janela**, que é a que teria impedido a rodada
  de 14/08 de rodar com uma janela já consumida.

Este módulo lê a pasta oficial do OneDrive — **somente leitura**, regra 15.8 —
e monta o histórico a partir do que de fato foi emitido em cada rodada.

O que dá e o que não dá
-----------------------

A estrutura da planilha mudou várias vezes ao longo de 2026, e a semente respeita
o que cada arquivo permite:

* **rodadas 9 em diante** — a coluna ``N°`` do Termo é o ``IdTituloVortx`` e não
  se repete entre sacados. Dá para travar título a título;
* **rodadas 4 a 8** — a mesma coluna repete o número entre pessoas diferentes
  (1.420 linhas para 883 números na 4ª). Ali ela **não** é identidade de título,
  e travar por ela barraria o título errado numa rodada futura. A semente
  registra a janela, a data, a contagem e o valor, e deixa ``titulos`` vazio,
  dizendo o motivo em ``motivo_sem_travas``;
* **rodadas 1 a 3** — os arquivos não têm aba de Termo. Fica só a data, deduzida
  do nome da pasta, para a linha do tempo não ter buraco silencioso.

A numeração das rodadas é um problema real
------------------------------------------

As três fontes de número discordam entre si nos arquivos de agosto::

    pasta                    nome do arquivo     INFORMAÇÕES!D9   Termo assinado
    12.Recompra_18_07_26     Leve_Saude_11       1  (obsoleto)    "FIDC Leve - 11"
    13.Recompra_10_08_2026   Leve_Saude_13       1  (obsoleto)    "..._12_v1.0"
    14.Recompra_17_08_2026   Leve_Saude_14       14               —

A pasta ``13`` traz arquivo numerado 13 e Termo assinado numerado 12; a ``14``
traz arquivo e ``D9`` numerados 14, enquanto a contagem de pastas daria 13.
Alguém pulou um número em agosto, e não há dado que diga quem.

A semente **não resolve isso em silêncio**. Vale o número do nome do arquivo —
é o que foi entregue e o que quem opera vê —, os três candidatos ficam gravados
em cada registro, e o conflito é devolvido em ``conflitos``. Como a sequência
resultante tem buraco, :attr:`Historico.numeracao_confiavel` fica falsa e a
interface passa a **exigir** o número da rodada em vez de sugerir.
"""

import re
from decimal import Decimal
from pathlib import Path

from core import config, leitura
from core import normalizacao as nz
from core.historico import Historico
from core.retroativo import (SITUACAO_COMPLETA, SITUACAO_SEM_ESTRUTURA,
                             SITUACAO_SO_TERMO, _escolher_arquivo,
                             abas_do_arquivo, ler_termo_emitido)

#  ``dd_mm_aa`` ou ``dd_mm_aaaa`` no nome da pasta: 14.Recompra_17_08_2026.
_DATA_NA_PASTA = re.compile(r"_(\d{2})_(\d{2})_(\d{2,4})\s*$")
_NUMERO_NO_NOME = re.compile(r"Leve_Saude_(\d+)_Recompra", re.IGNORECASE)


class Achado:
    """Uma rodada lida do OneDrive, com a procedência de cada campo."""

    def __init__(self, pasta):
        self.pasta = Path(pasta)
        self.arquivo = None
        self.abas = []
        self.situacao = SITUACAO_SEM_ESTRUTURA
        self.numero = None
        self.numeros_candidatos = {}
        self.data_recompra = None
        self.janela_inicio = None
        self.janela_fim = None
        self.origem_janela = ""
        self.linhas_termo = 0
        self.titulos = []
        self.id_confiavel = False
        self.motivo_sem_travas = ""
        self.valor_total = Decimal(0)
        self.juros_mora = None
        self.multa = None
        self.observacao = ""

    @property
    def conflito_de_numero(self):
        distintos = {v for v in self.numeros_candidatos.values() if v}
        return len(distintos) > 1

    def __repr__(self):
        return (f"<Achado {self.pasta.name} n={self.numero} "
                f"{self.linhas_termo} linhas no Termo>")


# ---------------------------------------------------------------------------
#  Leitura de uma pasta
# ---------------------------------------------------------------------------
def _data_da_pasta(pasta):
    """``14.Recompra_17_08_2026`` -> 17/08/2026."""
    achado = _DATA_NA_PASTA.search(pasta.name)
    if not achado:
        return None
    dia, mes, ano = achado.groups()
    ano = int(ano)
    if ano < 100:
        ano += 2000
    return nz.data(f"{dia}/{mes}/{ano}")


def _numero_da_pasta(pasta):
    achado = re.match(r"(\d+)\.", pasta.name)
    return int(achado.group(1)) - 1 if achado else None


def _ler_informacoes(caminho):
    """As células de parâmetro, quando a aba INFORMAÇÕES existe."""
    try:
        celulas = leitura.ler_celulas(
            caminho, config.ABA_INFORMACOES,
            [config.CEL_DATA_RECOMPRA, config.CEL_NUMERO_RODADA,
             config.CEL_JANELA_INICIO, config.CEL_JANELA_FIM,
             config.CEL_JUROS_MORA, config.CEL_MULTA],
        )
    except (KeyError, OSError, ValueError):
        return {}
    return celulas


def _inteiro(v):
    try:
        n = int(float(str(v).strip()))
    except (TypeError, ValueError):
        return None
    return n if n > 0 else None


def _decimal(v):
    try:
        return Decimal(str(v).strip())
    except (TypeError, ValueError, ArithmeticError):
        return None


def ler_pasta(pasta, log=None):
    """Levanta uma pasta de rodada. Nunca escreve nada nela."""
    log = log or (lambda *a, **k: None)
    achado = Achado(pasta)
    achado.data_recompra = _data_da_pasta(pasta)
    achado.numeros_candidatos["pasta"] = _numero_da_pasta(pasta)

    achado.arquivo = _escolher_arquivo(pasta)
    if achado.arquivo is None:
        achado.observacao = "nenhuma planilha na pasta"
        log(f"  {pasta.name}: sem planilha.", "aviso")
        return achado

    no_nome = _NUMERO_NO_NOME.search(achado.arquivo.name)
    achado.numeros_candidatos["arquivo"] = int(no_nome.group(1)) if no_nome else None

    achado.abas = abas_do_arquivo(achado.arquivo)
    nomes = {nz.sem_acento(a).strip() for a in achado.abas}
    if {"VORTX", "GRAFENO"} <= nomes and "TERMO DE RECOMPRA" in nomes:
        achado.situacao = SITUACAO_COMPLETA
    elif "TERMO DE RECOMPRA" in nomes:
        achado.situacao = SITUACAO_SO_TERMO

    # -- parâmetros da aba INFORMAÇÕES --------------------------------------
    if any("INFORMA" in nz.sem_acento(a) for a in achado.abas):
        celulas = _ler_informacoes(achado.arquivo)
        achado.numeros_candidatos["informacoes"] = _inteiro(
            celulas.get(config.CEL_NUMERO_RODADA))
        achado.data_recompra = (nz.data(celulas.get(config.CEL_DATA_RECOMPRA))
                                or achado.data_recompra)
        achado.janela_inicio = nz.data(celulas.get(config.CEL_JANELA_INICIO))
        achado.janela_fim = nz.data(celulas.get(config.CEL_JANELA_FIM))
        if achado.janela_inicio and achado.janela_fim:
            achado.origem_janela = "INFORMAÇÕES!D10:D11"
        achado.juros_mora = _decimal(celulas.get(config.CEL_JUROS_MORA))
        achado.multa = _decimal(celulas.get(config.CEL_MULTA))

    # -- o Termo que de fato saiu ------------------------------------------
    if achado.situacao == SITUACAO_SEM_ESTRUTURA:
        achado.observacao = (f"a planilha não tem aba de Termo "
                             f"(abas: {', '.join(achado.abas) or 'nenhuma'})")
        log(f"  {pasta.name}: {achado.observacao}", "aviso")
        return achado

    try:
        linhas, confiavel = ler_termo_emitido(achado.arquivo)
    except (KeyError, OSError, ValueError) as erro:
        achado.observacao = f"não consegui ler o Termo: {erro}"
        log(f"  {pasta.name}: {achado.observacao}", "aviso")
        return achado
    finally:
        leitura.fechar_cache()

    achado.linhas_termo = len(linhas)
    achado.id_confiavel = confiavel
    achado.valor_total = sum((l["valor_recompra"] or Decimal(0) for l in linhas),
                             Decimal(0))

    #  A trava só é semeada com identificador que é de fato identidade de
    #  título. Travar por número que se repete entre sacados barraria o título
    #  errado numa rodada futura — é a mesma regra do reprocessamento retroativo.
    if confiavel:
        achado.titulos = sorted({l["identificador"] for l in linhas
                                 if l["identificador"]})
    else:
        distintos = len({l["identificador"] for l in linhas if l["identificador"]})
        achado.motivo_sem_travas = (
            f"a coluna de número do Termo não identifica título nesta rodada "
            f"({achado.linhas_termo} linhas para {distintos} números, repetidos "
            f"entre sacados diferentes) — janela e valor foram registrados, "
            f"títulos não"
        )

    #  Sem janela na aba INFORMAÇÕES, o próprio Termo a revela: os vencimentos
    #  que entraram nele são, por definição, os da janela daquela rodada.
    if not (achado.janela_inicio and achado.janela_fim):
        vencimentos = [l["vencimento"] for l in linhas if l["vencimento"]]
        if vencimentos:
            achado.janela_inicio = min(vencimentos)
            achado.janela_fim = max(vencimentos)
            achado.origem_janela = "vencimentos do Termo emitido"

    return achado


# ---------------------------------------------------------------------------
#  Numeração
# ---------------------------------------------------------------------------
def _resolver_numero(achado):
    """
    Qual número vale. Ordem: nome do arquivo, INFORMAÇÕES!D9, contagem de pastas.

    O nome do arquivo vem primeiro porque é o que foi entregue e o que quem
    opera vê. ``D9`` ficou obsoleto em vários arquivos (marca 1 na rodada de
    10/08) e a contagem de pastas assume que nenhum número foi pulado — que é
    justamente o que aconteceu em agosto.
    """
    for fonte in ("arquivo", "informacoes", "pasta"):
        numero = achado.numeros_candidatos.get(fonte)
        if numero:
            return numero, fonte
    return None, ""


# ---------------------------------------------------------------------------
#  Semeadura
# ---------------------------------------------------------------------------
def semear(raiz=None, destino=None, log=None, sobrescrever=False):
    """
    Monta o ``historico_recompras.json`` a partir da pasta do OneDrive.

    Devolve um resumo para a tela. Recusa-se a apagar um histórico que já tenha
    rodadas, a menos que ``sobrescrever`` seja explícito: o histórico é a trava
    contra cobrar duas vezes, e sobrescrever por acidente a desarma.
    """
    log = log or (lambda *a, **k: None)
    raiz = Path(raiz or config.PASTA_ONEDRIVE)
    caminho = Path(destino or config.ARQUIVO_HISTORICO)
    if not raiz.exists():
        raise FileNotFoundError(
            f"Pasta das rodadas anteriores não encontrada:\n{raiz}\n"
            f"Confira o caminho ou se o OneDrive está sincronizado."
        )

    hist = Historico(caminho)
    if hist.rodadas and not sobrescrever:
        raise RuntimeError(
            f"O histórico já tem {len(hist.rodadas)} rodada(s). Semear de novo "
            f"reescreveria a trava contra recompra repetida. Marque a opção de "
            f"sobrescrever se é isso que você quer, ou renomeie "
            f"{caminho.name} antes."
        )

    log(f"Semeando o histórico a partir de {raiz}", "destaque")
    log("Leitura apenas: nada é escrito, movido ou renomeado no OneDrive.")

    achados = []
    for pasta in sorted(raiz.iterdir()):
        if not pasta.is_dir() or "recompra" not in pasta.name.lower():
            continue
        log(f"  {pasta.name}…")
        achados.append(ler_pasta(pasta, log=log))

    hist.rodadas = []
    conflitos = []
    sem_travas = []
    for achado in achados:
        numero, fonte = _resolver_numero(achado)
        achado.numero = numero
        if numero is None:
            log(f"  {achado.pasta.name}: sem número identificável — fora do "
                f"histórico.", "aviso")
            continue
        if achado.conflito_de_numero:
            conflitos.append({
                "pasta": achado.pasta.name,
                "escolhido": numero,
                "fonte": fonte,
                "candidatos": dict(achado.numeros_candidatos),
            })
        if achado.motivo_sem_travas:
            sem_travas.append({"numero": numero,
                               "motivo": achado.motivo_sem_travas})

        hist.registrar(
            numero=numero,
            data_recompra=achado.data_recompra,
            janela_inicio=achado.janela_inicio,
            janela_fim=achado.janela_fim,
            titulos=achado.titulos,
            valor_total=achado.valor_total,
            arquivo_saida=achado.arquivo.name if achado.arquivo else "",
            observacao=(achado.observacao
                        or "importado do Termo emitido pela semente"),
            juros_mora=achado.juros_mora,
            multa=achado.multa,
            gravar=False,
            extras={
                "pasta": achado.pasta.name,
                "situacao": achado.situacao,
                "linhas_termo": achado.linhas_termo,
                "origem_numero": fonte,
                "numeros_candidatos": dict(achado.numeros_candidatos),
                "origem_janela": achado.origem_janela,
                "id_confiavel": achado.id_confiavel,
                "motivo_sem_travas": achado.motivo_sem_travas,
                "semeado": True,
            },
        )
        valor = (f"{achado.valor_total:,.2f}"
                 .replace(",", "@").replace(".", ",").replace("@", "."))
        janela = (f"{nz.br(achado.janela_inicio)} a {nz.br(achado.janela_fim)}"
                  if achado.janela_inicio else "não identificada")
        log(f"    rodada {numero} (número pelo {fonte}): "
            f"{achado.linhas_termo} linhas no Termo · janela {janela} · "
            f"R$ {valor} · {len(achado.titulos)} títulos travados",
            "info" if achado.titulos else "aviso")

    hist.gravar()

    resumo = {
        "arquivo": str(caminho),
        "rodadas": len(hist.rodadas),
        "titulos_travados": len(hist.ids_recomprados),
        "conflitos": conflitos,
        "sem_travas": sem_travas,
        "numeracao_confiavel": hist.numeracao_confiavel,
        "proximo_numero": hist.proximo_numero(),
        "ultima": hist.resumo_ultima(),
        "detalhe": [
            {"numero": r.get("numero"), "pasta": r.get("pasta", ""),
             "janela": f"{nz.br(r.get('janela_inicio'))} a {nz.br(r.get('janela_fim'))}",
             "data_recompra": nz.br(r.get("data_recompra")),
             "linhas": r.get("linhas_termo", 0),
             "titulos": len(r.get("titulos") or []),
             "valor": str(r.get("valor_total") or "0"),
             "situacao": r.get("situacao", "")}
            for r in sorted(hist.rodadas, key=lambda x: x.get("numero") or 0)
        ],
    }

    log("")
    log(f"Histórico gravado: {len(hist.rodadas)} rodadas, "
        f"{len(hist.ids_recomprados)} títulos travados.", "ok")
    if conflitos:
        log(f"⚠ {len(conflitos)} rodada(s) com numeração conflitante entre a "
            f"pasta, o nome do arquivo e a aba INFORMAÇÕES:", "aviso")
        for c in conflitos:
            log(f"    {c['pasta']}: escolhido {c['escolhido']} (pelo "
                f"{c['fonte']}) entre {c['candidatos']}", "aviso")
        log("  Por isso a tela vai exigir o número da rodada, em vez de "
            "sugerir. Confira contra o Termo assinado.", "aviso")
    if sem_travas:
        log(f"⚠ {len(sem_travas)} rodada(s) sem trava título a título — a coluna "
            f"de número do Termo não identifica título naquele layout:", "aviso")
        for s in sem_travas:
            log(f"    rodada {s['numero']}: {s['motivo']}", "aviso")
    return resumo
