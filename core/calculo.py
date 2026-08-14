# -*- coding: utf-8 -*-
"""
Multa, juros e janela quinzenal.
================================

Juros e multa **sempre** vêm da aba INFORMAÇÕES do template (decisão 1.5 e
regra 15.3). Nada de taxa escrita no código: se a Leve Saúde renegociar o
contrato, a automação acompanha sozinha.

A conta reproduz exatamente as colunas AW a AZ da aba VORTX::

    prazo = (data_recompra - vencimento) em dias      # o dia do vencimento não conta
    multa = valor_nominal * taxa_multa
    juros = ((1 + taxa_juros_diaria) ** prazo - 1) * valor_nominal
    total = valor_nominal + multa + juros

Precisão total nos passos intermediários; arredonda só na apresentação.
"""

import calendar
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP

from core import config
from core import leitura
from core import normalizacao as nz

CENTAVO = Decimal("0.01")


# ---------------------------------------------------------------------------
#  Parâmetros lidos da aba INFORMAÇÕES
# ---------------------------------------------------------------------------
class Parametros:
    """Os números da rodada, do jeito que a planilha os guarda."""

    __slots__ = ("juros_mora", "multa", "data_recompra", "numero_rodada",
                 "janela_inicio", "janela_fim", "origem")

    def __init__(self, **campos):
        for nome in self.__slots__:
            setattr(self, nome, campos.get(nome))

    def __repr__(self):
        return (f"<Parametros rodada={self.numero_rodada} "
                f"recompra={self.data_recompra} juros={self.juros_mora} "
                f"multa={self.multa}>")


def ler_parametros(caminho_template, log=None):
    """Lê juros, multa, datas e número da rodada da aba INFORMAÇÕES."""
    log = log or (lambda *a, **k: None)
    celulas = leitura.ler_celulas(
        caminho_template, config.ABA_INFORMACOES,
        [config.CEL_DATA_RECOMPRA, config.CEL_NUMERO_RODADA,
         config.CEL_JANELA_INICIO, config.CEL_JANELA_FIM,
         config.CEL_JUROS_MORA, config.CEL_MULTA],
    )
    juros = _taxa(celulas.get(config.CEL_JUROS_MORA))
    multa = _taxa(celulas.get(config.CEL_MULTA))
    if juros is None or multa is None:
        raise ValueError(
            f"Não consegui ler JUROS MORA ({config.CEL_JUROS_MORA}) e MULTA "
            f"({config.CEL_MULTA}) da aba {config.ABA_INFORMACOES}. "
            f"Sem esses dois números o cálculo não roda (decisão 1.5)."
        )

    numero = celulas.get(config.CEL_NUMERO_RODADA)
    parametros = Parametros(
        juros_mora=juros,
        multa=multa,
        data_recompra=nz.data(celulas.get(config.CEL_DATA_RECOMPRA)),
        numero_rodada=int(numero) if isinstance(numero, (int, float)) else None,
        janela_inicio=nz.data(celulas.get(config.CEL_JANELA_INICIO)),
        janela_fim=nz.data(celulas.get(config.CEL_JANELA_FIM)),
        origem=str(caminho_template),
    )
    log(f"  INFORMAÇÕES: juros {juros * 100:.6f}% a.d., multa {multa * 100:.2f}%.")
    return parametros


def _taxa(v):
    if v is None or isinstance(v, str) and not v.strip():
        return None
    if isinstance(v, str):
        v = v.strip().replace("%", "").replace(",", ".")
        try:
            return Decimal(v) / (Decimal(100) if float(v) > 1 else Decimal(1))
        except (ValueError, ArithmeticError):
            return None
    try:
        return Decimal(str(v))
    except ArithmeticError:
        return None


# ---------------------------------------------------------------------------
#  Cálculo do título
# ---------------------------------------------------------------------------
class Calculo:
    """Valores calculados de um título do Termo."""

    __slots__ = ("prazo", "valor_nominal", "multa", "juros", "total")

    def __init__(self, prazo, valor_nominal, multa, juros, total):
        self.prazo = prazo
        self.valor_nominal = valor_nominal
        self.multa = multa
        self.juros = juros
        self.total = total


def calcular(valor_nominal, vencimento, data_recompra, taxa_juros, taxa_multa):
    """Multa e juros de um título, na conta da coluna AZ da aba VORTX."""
    if valor_nominal is None or vencimento is None or data_recompra is None:
        raise ValueError("Valor, vencimento e data de recompra são obrigatórios.")

    prazo = (data_recompra - vencimento).days
    if prazo < 0:
        prazo = 0
    valor = Decimal(valor_nominal)
    multa = valor * Decimal(taxa_multa)
    juros = ((Decimal(1) + Decimal(taxa_juros)) ** prazo - Decimal(1)) * valor
    return Calculo(prazo, valor, multa, juros, valor + multa + juros)


def arredondar(v):
    """Duas casas, meio para cima — só na apresentação."""
    if v is None:
        return None
    return Decimal(v).quantize(CENTAVO, rounding=ROUND_HALF_UP)


def somar(valores):
    total = Decimal(0)
    for v in valores:
        if v is not None:
            total += Decimal(v)
    return total


# ---------------------------------------------------------------------------
#  Janela quinzenal (decisão 1.4 e seção 8.5)
# ---------------------------------------------------------------------------
def sugerir_janela(data_recompra):
    """
    Janela quinzenal estrita, deduzida da data de recompra.

    * recompra na 1ª quinzena  -> dia 16 ao último dia do **mês anterior**;
    * recompra na 2ª quinzena  -> dia 1 ao 15 do **mês corrente**.
    """
    data_recompra = nz.data(data_recompra)
    if data_recompra is None:
        return None, None
    if data_recompra.day <= 15:
        ano, mes = (data_recompra.year, data_recompra.month - 1)
        if mes == 0:
            ano, mes = ano - 1, 12
        return date(ano, mes, 16), date(ano, mes, calendar.monthrange(ano, mes)[1])
    return (date(data_recompra.year, data_recompra.month, 1),
            date(data_recompra.year, data_recompra.month, 15))


def janela_valida(inicio, fim):
    """True quando a janela segue o padrão dia 1–15 ou 16–fim do mês."""
    inicio, fim = nz.data(inicio), nz.data(fim)
    if not inicio or not fim or inicio > fim:
        return False
    if inicio.year != fim.year or inicio.month != fim.month:
        return False
    ultimo = calendar.monthrange(inicio.year, inicio.month)[1]
    return (inicio.day, fim.day) in ((1, 15), (16, ultimo))


def janela_seguinte(fim_anterior):
    """A janela que vem logo depois da que terminou em ``fim_anterior``."""
    fim_anterior = nz.data(fim_anterior)
    if fim_anterior is None:
        return None, None
    inicio = fim_anterior + timedelta(days=1)
    if inicio.day == 1:
        return inicio, date(inicio.year, inicio.month, 15)
    ultimo = calendar.monthrange(inicio.year, inicio.month)[1]
    return inicio, date(inicio.year, inicio.month, ultimo)


# ---------------------------------------------------------------------------
#  Dias úteis e o custo da defasagem (seção 2.4)
# ---------------------------------------------------------------------------
def _pascoa(ano):
    """Domingo de Páscoa pelo algoritmo de Meeus/Jones/Butcher."""
    a = ano % 19
    b, c = divmod(ano, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes, dia = divmod(h + l - 7 * m + 114, 31)
    return date(ano, mes, dia + 1)


def feriados_nacionais(ano):
    """Feriados nacionais fixos e os móveis derivados da Páscoa."""
    pascoa = _pascoa(ano)
    return {
        date(ano, 1, 1), date(ano, 4, 21), date(ano, 5, 1), date(ano, 9, 7),
        date(ano, 10, 12), date(ano, 11, 2), date(ano, 11, 15),
        date(ano, 11, 20), date(ano, 12, 25),
        pascoa - timedelta(days=48),      # segunda de carnaval
        pascoa - timedelta(days=47),      # terça de carnaval
        pascoa - timedelta(days=2),       # sexta-feira santa
        pascoa + timedelta(days=60),      # corpus christi
    }


def dias_uteis(inicio, fim):
    """
    Dias úteis entre duas datas, sem contar o dia inicial.

    Sexta 07/08 -> segunda 10/08 dá **um** dia útil. Foi essa a defasagem da
    12ª rodada, e nesse único dia útil 20 sacados pagaram o próprio boleto.
    """
    inicio, fim = nz.data(inicio), nz.data(fim)
    if not inicio or not fim or fim <= inicio:
        return 0
    feriados = set()
    for ano in range(inicio.year, fim.year + 1):
        feriados |= feriados_nacionais(ano)
    total = 0
    dia = inicio + timedelta(days=1)
    while dia <= fim:
        if dia.weekday() < 5 and dia not in feriados:
            total += 1
        dia += timedelta(days=1)
    return total


def custo_defasagem(dias):
    """
    Estimativa do que a defasagem custa, para a interface mostrar na hora.

    O modelo saiu da própria 12ª rodada: entre a extração de 03/08 e a de
    07/08 saíram 74 títulos do Termo, R$ 89.640,46, em quatro dias úteis.
    Um dia útil previa ~20 títulos e foram exatamente 20.
    """
    if dias <= 0:
        return {"dias": 0, "titulos_min": 0, "titulos_max": 0, "valor": Decimal(0)}
    return {
        "dias": dias,
        "titulos_min": 18 * dias,
        "titulos_max": 20 * dias,
        "valor": config.VALOR_POR_DIA_UTIL * dias,
    }


def frase_defasagem(dias):
    """A frase que vai para a tela, no formato pedido na seção 2.4."""
    if dias <= 0:
        return "Extração no dia da recompra — defasagem zero."
    c = custo_defasagem(dias)
    rotulo = "1 dia útil" if dias == 1 else f"{dias} dias úteis"
    valor = f"{c['valor']:,.0f}".replace(",", ".")
    return (
        f"Extração com {rotulo} de defasagem — estimativa de "
        f"{c['titulos_min']} a {c['titulos_max']} títulos que podem ser pagos "
        f"antes da recompra, cerca de R$ {valor}."
    )
